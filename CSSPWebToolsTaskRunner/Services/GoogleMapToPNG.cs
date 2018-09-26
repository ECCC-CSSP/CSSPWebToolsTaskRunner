using CSSPEnumsDLL.Enums;
using CSSPModelsDLL.Models;
using CSSPWebToolsDBDLL.Services;
using CSSPWebToolsTaskRunner.Services.Resources;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Net;

namespace CSSPWebToolsTaskRunner.Services
{
    public partial class GoogleMapToPNG
    {
        #region Variables
        string HideVerticalScale = "";
        string HideHorizontalScale = "";
        string HideNorthArrow = "";
        string HideSubsectorName = "";
        int IconSize = 10;
        #endregion Variables

        #region Properties
        private TaskRunnerBaseService _TaskRunnerBaseService { get; set; }
        private string _Parameters { get; set; }
        private TVItemService _TVItemService { get; set; }
        private MapInfoService _MapInfoService { get; set; }
        private MapInfoPointService _MapInfoPointService { get; set; }
        private PolSourceSiteService _PolSourceSiteService { get; set; }
        private PolSourceObservationService _PolSourceObservationService { get; set; }
        private PolSourceObservationIssueService _PolSourceObservationIssueService { get; set; }
        private TVFileService _TVFileService { get; set; }

        public string DirName { get; set; }
        private string FileNameNW { get; set; }
        private string FileNameNE { get; set; }
        private string FileNameSW { get; set; }
        private string FileNameSE { get; set; }
        private string FileNameFull { get; set; }
        private string FileNameInset { get; set; }
        private string FileNameInsetFinal { get; set; }
        public string FileNameExtra { get; set; }
        private int GoogleImageWidth { get; set; }
        private int GoogleImageHeight { get; set; }
        private int GoogleLogoHeight { get; set; }
        private LanguageEnum LanguageRequest { get; set; }
        private string MapType { get; set; }
        private int Zoom { get; set; }
        public int TVFileTVItemID { get; set; }
        public string FileNameFullAnnotated { get; set; }
        #endregion Properties

        #region Constructors
        public GoogleMapToPNG(TaskRunnerBaseService taskRunnerBaseService, string HideVerticalScale, string HideHorizontalScale, string HideNorthArrow, string HideSubsectorName)
        {
            _TaskRunnerBaseService = taskRunnerBaseService;
            this.HideVerticalScale = HideVerticalScale;
            this.HideHorizontalScale = HideHorizontalScale;
            this.HideNorthArrow = HideNorthArrow;
            this.HideSubsectorName = HideSubsectorName;
            _TVItemService = new TVItemService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            _MapInfoService = new MapInfoService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            _MapInfoPointService = new MapInfoPointService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            _PolSourceSiteService = new PolSourceSiteService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            _PolSourceObservationService = new PolSourceObservationService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            _PolSourceObservationIssueService = new PolSourceObservationIssueService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
            _TVFileService = new TVFileService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);

            LoadDefaults();
        }
        public GoogleMapToPNG()
        {
        }

        #endregion Constructors

        #region Functions public
        public bool CreateSubsectorGoogleMapPNGForPolSourceSites(int SubsectorTVItemID, string MapType)
        {
            //string NotUsed = "";
            this.MapType = MapType;

            TVItemModel tvItemModelSubsector = _TVItemService.GetTVItemModelWithTVItemIDDB(SubsectorTVItemID);
            if (!string.IsNullOrWhiteSpace(tvItemModelSubsector.Error))
            {
                // need to set error
                return false;
            }

            List<MapInfoPointModel> mapInfoPointModelSubsectorList = _MapInfoPointService.GetMapInfoPointModelListWithTVItemIDAndTVTypeAndMapInfoDrawTypeDB(SubsectorTVItemID, TVTypeEnum.Subsector, MapInfoDrawTypeEnum.Polygon);
            if (mapInfoPointModelSubsectorList.Count == 0)
            {
                // need to set the error
                return false;
            }

            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 10);

            double MaxLat = mapInfoPointModelSubsectorList.Select(c => c.Lat).Max();
            double MinLat = mapInfoPointModelSubsectorList.Select(c => c.Lat).Min();
            double MaxLng = mapInfoPointModelSubsectorList.Select(c => c.Lng).Max();
            double MinLng = mapInfoPointModelSubsectorList.Select(c => c.Lng).Min();


            List<MapInfoPointModel> mapInfoPointModelPolSourceSiteList = _MapInfoPointService.GetMapInfoPointModelListWithParentIDAndTVTypeAndMapInfoDrawTypeDB(SubsectorTVItemID, TVTypeEnum.PolSourceSite, MapInfoDrawTypeEnum.Point);
            List<TVItemModel> tvItemModelPolSourceSiteList = _TVItemService.GetChildrenTVItemModelListWithTVItemIDAndTVTypeDB(SubsectorTVItemID, TVTypeEnum.PolSourceSite).Where(c => c.IsActive == true).ToList();

            mapInfoPointModelPolSourceSiteList = (from c in mapInfoPointModelPolSourceSiteList
                                                  from t in tvItemModelPolSourceSiteList // all active sites
                                                  where c.TVItemID == t.TVItemID
                                                  select c).ToList();

            if (mapInfoPointModelPolSourceSiteList.Count > 0)
            {
                MaxLat = mapInfoPointModelPolSourceSiteList.Select(c => c.Lat).Max();
                MinLat = mapInfoPointModelPolSourceSiteList.Select(c => c.Lat).Min();
                MaxLng = mapInfoPointModelPolSourceSiteList.Select(c => c.Lng).Max();
                MinLng = mapInfoPointModelPolSourceSiteList.Select(c => c.Lng).Min();
            }

            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 15);

            CoordMap coordMap = GetMapCoordinateWhileGettingGooglePNG(MinLat, MaxLat, MinLng, MaxLng);

            //int a = 234;

            //CoordMap coordMap = new CoordMap()
            //{
            //    NorthEast = new Coord() { Lat = 46.5364151f, Lng = -64.55215f, Ordinal = 0 },
            //    SouthWest = new Coord() { Lat = 46.23907f, Lng = -64.99161f, Ordinal = 0 },
            //};

            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 20);

            if (!DrawSubsectorPolSourceSites(coordMap, tvItemModelSubsector, mapInfoPointModelSubsectorList, mapInfoPointModelPolSourceSiteList, tvItemModelPolSourceSiteList))
            {
                return false;
            }

            return true;
        }
        public bool CreateSubsectorGoogleMapPNGForMWQMSites(int SubsectorTVItemID, string MapType)
        {
            string NotUsed = "";
            this.MapType = MapType;

            TVItemModel tvItemModelSubsector = _TVItemService.GetTVItemModelWithTVItemIDDB(SubsectorTVItemID);
            if (!string.IsNullOrWhiteSpace(tvItemModelSubsector.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes._IsRequired, TaskRunnerServiceRes.SubsectorTVItemID);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_IsRequired", TaskRunnerServiceRes.SubsectorTVItemID);
                return false;
            }

            List<MapInfoPointModel> mapInfoPointModelSubsectorList = _MapInfoPointService.GetMapInfoPointModelListWithTVItemIDAndTVTypeAndMapInfoDrawTypeDB(SubsectorTVItemID, TVTypeEnum.Subsector, MapInfoDrawTypeEnum.Polygon);
            if (mapInfoPointModelSubsectorList.Count == 0)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFindPolygonForSubsector_, tvItemModelSubsector.TVText);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotFindPolygonForSubsector_", tvItemModelSubsector.TVText);
                return false;
            }

            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 10);

            double MaxLat = mapInfoPointModelSubsectorList.Select(c => c.Lat).Max();
            double MinLat = mapInfoPointModelSubsectorList.Select(c => c.Lat).Min();
            double MaxLng = mapInfoPointModelSubsectorList.Select(c => c.Lng).Max();
            double MinLng = mapInfoPointModelSubsectorList.Select(c => c.Lng).Min();

            List<MapInfoPointModel> mapInfoPointModelMWQMSiteList = _MapInfoPointService.GetMapInfoPointModelListWithParentIDAndTVTypeAndMapInfoDrawTypeDB(SubsectorTVItemID, TVTypeEnum.MWQMSite, MapInfoDrawTypeEnum.Point);
            List<TVItemModel> tvItemModelMWQMSiteList = _TVItemService.GetChildrenTVItemModelListWithTVItemIDAndTVTypeDB(SubsectorTVItemID, TVTypeEnum.MWQMSite).Where(c => c.IsActive == true).ToList();

            mapInfoPointModelMWQMSiteList = (from c in mapInfoPointModelMWQMSiteList
                                             from t in tvItemModelMWQMSiteList // all active sites
                                             where c.TVItemID == t.TVItemID
                                             select c).ToList();


            if (mapInfoPointModelMWQMSiteList.Count > 0)
            {
                MaxLat = mapInfoPointModelMWQMSiteList.Select(c => c.Lat).Max();
                MinLat = mapInfoPointModelMWQMSiteList.Select(c => c.Lat).Min();
                MaxLng = mapInfoPointModelMWQMSiteList.Select(c => c.Lng).Max();
                MinLng = mapInfoPointModelMWQMSiteList.Select(c => c.Lng).Min();
            }

            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 15);

            CoordMap coordMap = GetMapCoordinateWhileGettingGooglePNG(MinLat, MaxLat, MinLng, MaxLng);

            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 20);

            if (!DrawSubsectorMWQMSites(coordMap, tvItemModelSubsector, mapInfoPointModelSubsectorList, mapInfoPointModelMWQMSiteList, tvItemModelMWQMSiteList))
            {
                return false;
            }

            return true;
        }
        #endregion Functions public

        #region Functions private

        #region Functions private draw icons
        private void DrawAgricultureIcon(Graphics g, Pen pen, SolidBrush brush, int width, int LatY, int LngX)
        {
            Point[] pointArr = new List<Point>()
                            {
                                new Point() { X = (int)LngX - (width/3), Y = (int)LatY + (width/2) },
                                new Point() { X = (int)LngX + (width/3), Y = (int)LatY + (width/2) },
                                new Point() { X = (int)LngX + (width/3), Y = (int)LatY - (width/2) },
                                new Point() { X = (int)LngX - (width/3), Y = (int)LatY - (width/2) },
                                new Point() { X = (int)LngX - (width/3), Y = (int)LatY + (width/2) },
                            }.ToArray();
            g.DrawPolygon(pen, pointArr);
            g.FillPolygon(brush, pointArr);
        }
        private void DrawForestedIcon(Graphics g, Pen pen, SolidBrush brush, int width, int LatY, int LngX)
        {
            Point[] pointArr = new List<Point>()
                            {
                                new Point() { X = (int)LngX - (width/2), Y = (int)LatY + (width/2) },
                                new Point() { X = (int)LngX, Y = (int)LatY + (width/2) },
                                new Point() { X = (int)LngX, Y = (int)LatY + (width/2)*2 },
                                new Point() { X = (int)LngX, Y = (int)LatY + (width/2) },
                                new Point() { X = (int)LngX + (width/2), Y = (int)LatY + (width/2) },
                                new Point() { X = (int)LngX, Y = (int)LatY - (width/2) },
                                new Point() { X = (int)LngX - (width/2), Y = (int)LatY + (width/2) },
                            }.ToArray();
            g.DrawPolygon(pen, pointArr);
            g.FillPolygon(brush, pointArr);
        }
        private void DrawIndustryIcon(Graphics g, Pen pen, SolidBrush brush, int width, int LatY, int LngX)
        {
            Point[] pointArr = new List<Point>()
                            {
                                new Point() { X = (int)LngX - (width/2), Y = (int)LatY + (width/2) },
                                new Point() { X = (int)LngX + (width/2), Y = (int)LatY + (width/2) },
                                new Point() { X = (int)LngX + (width/2), Y = (int)LatY },
                                new Point() { X = (int)LngX + (width/3), Y = (int)LatY },
                                new Point() { X = (int)LngX + (width/3), Y = (int)LatY - (width/2) },
                                new Point() { X = (int)LngX - (width/3), Y = (int)LatY - (width/2) },
                                new Point() { X = (int)LngX - (width/3), Y = (int)LatY },
                                new Point() { X = (int)LngX - (width/2), Y = (int)LatY },
                                new Point() { X = (int)LngX - (width/2), Y = (int)LatY + (width/2) },
                            }.ToArray();
            g.DrawPolygon(pen, pointArr);
            g.FillPolygon(brush, pointArr);
        }
        private void DrawMarineIcon(Graphics g, Pen pen, SolidBrush brush, int width, int LatY, int LngX)
        {
            Point[] pointArr = new List<Point>()
                            {
                                new Point() { X = (int)LngX - (width/2), Y = (int)LatY + (width/2) },
                                new Point() { X = (int)LngX - (width/2), Y = (int)LatY - (width/2) },
                                new Point() { X = (int)LngX, Y = (int)LatY - (width/2) },
                                new Point() { X = (int)LngX + (width/2), Y = (int)LatY },
                                new Point() { X = (int)LngX, Y = (int)LatY + (width/2) },
                                new Point() { X = (int)LngX - (width/2), Y = (int)LatY + (width/2) },
                            }.ToArray();
            g.DrawPolygon(pen, pointArr);
            g.FillPolygon(brush, pointArr);
        }
        private void DrawRecreationIcon(Graphics g, Pen pen, SolidBrush brush, int width, int LatY, int LngX)
        {
            Point[] pointArr = new List<Point>()
                            {
                                new Point() { X = (int)LngX - (width/2), Y = (int)LatY + (width/2) },
                                new Point() { X = (int)LngX + (width/2), Y = (int)LatY + (width/2) },
                                new Point() { X = (int)LngX + (width/2), Y = (int)LatY },
                                new Point() { X = (int)LngX, Y = (int)LatY - (width/2) },
                                new Point() { X = (int)LngX - (width/2), Y = (int)LatY },
                                new Point() { X = (int)LngX - (width/2), Y = (int)LatY + (width/2) },
                            }.ToArray();
            g.DrawPolygon(pen, pointArr);
            g.FillPolygon(brush, pointArr);
        }
        private void DrawUrbanIcon(Graphics g, Pen pen, SolidBrush brush, int width, int LatY, int LngX)
        {
            Point[] pointArr = new List<Point>()
                            {
                                new Point() { X = (int)LngX - (width/2), Y = (int)LatY + (width/2) },
                                new Point() { X = (int)LngX + (width/2), Y = (int)LatY + (width/2) },
                                new Point() { X = (int)LngX + (width/2), Y = (int)LatY - (width/3) },
                                new Point() { X = (int)LngX + (width/2), Y = (int)LatY + (width/2) },
                                new Point() { X = (int)LngX + (width/3), Y = (int)LatY + (width/2) },
                                new Point() { X = (int)LngX + (width/3), Y = (int)LatY - (width/4) },
                                new Point() { X = (int)LngX + (width/3), Y = (int)LatY + (width/2) },
                                new Point() { X = (int)LngX, Y = (int)LatY - (width/2) },
                                new Point() { X = (int)LngX, Y = (int)LatY + (width/2) },
                                new Point() { X = (int)LngX - (width/3), Y = (int)LatY + (width/2) },
                                new Point() { X = (int)LngX - (width/3), Y = (int)LatY - width },
                                new Point() { X = (int)LngX - (width/3), Y = (int)LatY + (width/2) },
                                new Point() { X = (int)LngX - (width/2), Y = (int)LatY + (width/2) },
                            }.ToArray();
            g.DrawPolygon(pen, pointArr);
            g.FillPolygon(brush, pointArr);
        }
        private void DrawOtherIcon(Graphics g, Pen pen, SolidBrush brush, int width, int LatY, int LngX)
        {
            g.DrawEllipse(pen, (int)LngX - (width / 2), (int)LatY - (width / 2), width, width);
            g.FillEllipse(brush, (int)LngX - (width / 2), (int)LatY - (width / 2), width, width);
        }
        #endregion Functions private draw icons

        private bool CombineAllImageIntoOne()
        {
            Rectangle cropRect = new Rectangle(0, 0, GoogleImageWidth, GoogleImageHeight - GoogleLogoHeight);

            using (Bitmap targetNW = new Bitmap(cropRect.Width, cropRect.Height))
            {
                using (Graphics g = Graphics.FromImage(targetNW))
                {
                    using (Bitmap srcNW = new Bitmap(DirName + FileNameNW))
                    {
                        g.DrawImage(srcNW, new Rectangle(0, 0, targetNW.Width, targetNW.Height), cropRect, GraphicsUnit.Pixel);
                    }
                }

                using (Bitmap targetNE = new Bitmap(cropRect.Width, cropRect.Height))
                {
                    using (Graphics g = Graphics.FromImage(targetNE))
                    {
                        using (Bitmap srcNE = new Bitmap(DirName + FileNameNE))
                        {
                            g.DrawImage(srcNE, new Rectangle(0, 0, targetNE.Width, targetNE.Height), cropRect, GraphicsUnit.Pixel);
                        }
                    }

                    using (Bitmap targetAll = new Bitmap(cropRect.Width * 2, cropRect.Height + GoogleImageHeight))
                    {
                        using (Graphics g = Graphics.FromImage(targetAll))
                        {
                            g.DrawImage(targetNW, new Point(0, 0));
                            g.DrawImage(targetNE, new Point(GoogleImageWidth, 0));
                            using (Bitmap srcSW = new Bitmap(DirName + FileNameSW))
                            {
                                g.DrawImage(srcSW, new Point(0, GoogleImageHeight - GoogleLogoHeight));
                            }
                            using (Bitmap srcSE = new Bitmap(DirName + FileNameSE))
                            {
                                g.DrawImage(srcSE, new Point(GoogleImageWidth, GoogleImageHeight - GoogleLogoHeight));
                            }

                            targetAll.Save(DirName + FileNameFull, ImageFormat.Png);
                        }
                    }
                }
            }

            return true;
        }
        private bool DeleteTempGoogleImageFiles()
        {
            try
            {
                FileInfo fi = new FileInfo(DirName + FileNameNW);
                if (fi.Exists)
                {
                    fi.Delete();
                }
                fi = new FileInfo(DirName + FileNameNE);
                if (fi.Exists)
                {
                    fi.Delete();
                }
                fi = new FileInfo(DirName + FileNameSW);
                if (fi.Exists)
                {
                    fi.Delete();
                }
                fi = new FileInfo(DirName + FileNameSE);
                if (fi.Exists)
                {
                    fi.Delete();
                }
            }
            catch (Exception)
            {
                //richTextBoxStatus.Text = ex.Message; // + ex.InnerException != null ? " Inner: " + ex.InnerException.Message : "";
            }

            return true;
        }
        private bool DrawHorizontalScale(Graphics g, CoordMap coordMap)
        {
            string NotUsed = "";
            if (string.IsNullOrWhiteSpace(HideHorizontalScale))
            {
                try
                {
                    int MinLngX = (int)(GoogleImageWidth * 2 * 3 / 5);
                    int MaxLngX = (int)(GoogleImageWidth * 2 * 4 / 5);

                    g.DrawLine(new Pen(Color.LightBlue, 2), MinLngX, GoogleImageHeight * 2 - GoogleLogoHeight - 40, MaxLngX, GoogleImageHeight * 2 - GoogleLogoHeight - 40);

                    double MinLng = coordMap.NorthEast.Lng - (coordMap.NorthEast.Lng - coordMap.SouthWest.Lng) / 5;
                    double distLng = _MapInfoService.CalculateDistance(coordMap.NorthEast.Lat * _MapInfoService.d2r, coordMap.NorthEast.Lng * _MapInfoService.d2r, coordMap.NorthEast.Lat * _MapInfoService.d2r, MinLng * _MapInfoService.d2r, _MapInfoService.R) / 1000;


                    string distText = distLng.ToString("F2") + " km";
                    Font font = new Font("Arial", 10, FontStyle.Regular);
                    Brush brush = new SolidBrush(Color.LightBlue);

                    SizeF sizeF = g.MeasureString(distText, font);

                    g.DrawString(distText, font, brush, GoogleImageWidth * 2 * 4 / 5, GoogleImageHeight * 2 - GoogleLogoHeight - 60);

                    font = new Font("Arial", 10, FontStyle.Regular);
                    brush = new SolidBrush(Color.LightBlue);

                    for (int i = 0; i < 10; i++)
                    {
                        if ((double)i > distLng)
                        {
                            g.DrawLine(new Pen(Color.LightBlue, 1), MaxLngX, GoogleImageHeight * 2 - GoogleLogoHeight - 40 - 1, MaxLngX, GoogleImageHeight * 2 - GoogleLogoHeight - 40 + 10 - 1);
                            break;
                        }
                        g.DrawLine(new Pen(Color.LightBlue, 1), MinLngX + (int)(i / distLng * (MaxLngX - MinLngX)), GoogleImageHeight * 2 - GoogleLogoHeight - 40 - 1, MinLngX + (int)(i / distLng * (MaxLngX - MinLngX)), GoogleImageHeight * 2 - GoogleLogoHeight - 40 + 10 - 1);

                        distText = i.ToString();

                        sizeF = g.MeasureString(distText, font);

                        g.DrawString(distText, font, brush, MinLngX + (int)(i / distLng * (MaxLngX - MinLngX)) - (sizeF.Width / 2), GoogleImageHeight * 2 - GoogleLogoHeight - 60);

                    }

                }
                catch (Exception ex)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotCreate_ImageError_, TaskRunnerServiceRes.Annotated, ex.Message + ex.InnerException != null ? " Inner: " + ex.InnerException.Message : "");
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotCreate_ImageError_", TaskRunnerServiceRes.Annotated, ex.Message + ex.InnerException != null ? " Inner: " + ex.InnerException.Message : "");
                    return false;
                }
            }

            return true;
        }
        private bool DrawImageBorder(Graphics g)
        {
            string NotUsed = "";
            try
            {
                g.DrawRectangle(new Pen(Color.Black, 6.0f), 0, 0, GoogleImageWidth * 2, GoogleImageHeight * 2 - GoogleLogoHeight);
            }
            catch (Exception ex)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotCreate_ImageError_, TaskRunnerServiceRes.Annotated, ex.Message + ex.InnerException != null ? " Inner: " + ex.InnerException.Message : "");
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotCreate_ImageError_", TaskRunnerServiceRes.Annotated, ex.Message + ex.InnerException != null ? " Inner: " + ex.InnerException.Message : "");
                return false;
            }

            return true;
        }
        private bool DrawLegendMWQMSites(Graphics g, CoordMap coordMap, List<MapInfoPointModel> mapInfoPointModelMWQMSiteList, List<TVItemModel> tvItemModelMWQMSiteList)
        {
            string NotUsed = "";
            try
            {
                int StartingHeight = 100;
                int CurrentHeight = 0;
                int CenterOfLegend = 0;
                Brush brush = new SolidBrush(Color.White);
                Pen pen = new Pen(Color.LightGreen, 6.0f);
                Font font = new Font("Arial", 16, FontStyle.Bold);

                int LegendHeight = 0;
                int LegendWidth = 200;
                // Calculating Total legend height
                string LegendText = "Legend";
                SizeF sizeF = g.MeasureString(LegendText, font);
                LegendHeight = LegendHeight + (int)(sizeF.Height) + 10;

                font = new Font("Arial", 12, FontStyle.Bold);
                string ApprovedText = "Approved"; // Approved, Restricted, Conditional Approved, Conditionaly Restricted, Prohibited, Unclassified
                sizeF = g.MeasureString(ApprovedText, font);
                LegendHeight = LegendHeight + ((int)(sizeF.Height) + 5) * 12; // Passing, Fail, No Depuration, Colors (3)

                g.DrawRectangle(pen, GoogleImageWidth * 2 - LegendWidth, StartingHeight, LegendWidth - 5, LegendHeight);
                g.FillRectangle(brush, GoogleImageWidth * 2 - LegendWidth, StartingHeight, LegendWidth - 5, LegendHeight);

                CenterOfLegend = (GoogleImageWidth * 2) - LegendWidth / 2;
                brush = new SolidBrush(Color.LightBlue);

                // draw Legend title
                font = new Font("Arial", 16, FontStyle.Bold);
                sizeF = g.MeasureString(LegendText, font);
                CurrentHeight = StartingHeight + 10;
                g.DrawString(LegendText, font, brush, CenterOfLegend - (sizeF.Width / 2), CurrentHeight);

                CurrentHeight += 20;
                font = new Font("Arial", 8, FontStyle.Bold);
                brush = new SolidBrush(Color.Blue);
                List<string> ClassList = new List<string>()
                {
                    "Approved", "Restricted", "Conditional Approved", "Conditionaly Restricted", "Prohibited", "Unclassified"
                };
                List<string> ClassInitList = new List<string>()
                {
                    "A", "R", "CA", "CR", "P", "U"
                };

                sizeF = g.MeasureString(ClassInitList[2], font);
                int RectWidth = 24;
                int RectHeight = (int)(sizeF.Height) + 4;

                for (int i = 0, count = ClassList.Count; i < count; i++)
                {
                    brush = new SolidBrush(Color.Black);
                    sizeF = g.MeasureString(ClassList[i], font);
                    CurrentHeight += (int)(sizeF.Height) + 10;
                    g.DrawString(ClassList[i], font, brush, CenterOfLegend - LegendWidth / 2 + 50, CurrentHeight);

                    switch (i)
                    {
                        case 0: // Approved
                            {
                                brush = new SolidBrush(Color.Green);
                            }
                            break;
                        case 1: // Restricted
                            {
                                brush = new SolidBrush(Color.Red);
                            }
                            break;
                        case 2: // Conditional Approved
                            {
                                brush = new SolidBrush(Color.LightGreen);
                            }
                            break;
                        case 3: // Conditionaly Restricted
                            {
                                brush = new SolidBrush(Color.LightPink);
                            }
                            break;
                        case 4: // Prohibited
                            {
                                brush = new SolidBrush(Color.Black);
                            }
                            break;
                        case 5: // Unclassified
                            {
                                brush = new SolidBrush(Color.White);
                            }
                            break;
                        default:
                            break;
                    }
                    g.FillRectangle(brush, CenterOfLegend - LegendWidth / 2 + 20 - 2, CurrentHeight - 2, RectWidth, RectHeight);
                    g.DrawRectangle(new Pen(Color.Black, 2.0f), CenterOfLegend - LegendWidth / 2 + 20 - 2, CurrentHeight - 2, RectWidth, RectHeight);

                    if (i == 4)
                    {
                        brush = new SolidBrush(Color.White);
                    }
                    else if (i == 5)
                    {
                        brush = new SolidBrush(Color.Black);
                    }
                    else
                    {
                        brush = new SolidBrush(Color.Black);
                    }
                    g.DrawString(ClassInitList[i], font, brush, CenterOfLegend - LegendWidth / 2 + 20, CurrentHeight);


                }

                CurrentHeight += 20;

                List<string> ColorList = new List<string>()
                {
                    "Passed", "Failed", "No depuration"
                };

                List<string> FCCalcList = new List<string>() { "A", "B", "C", "D", "E", "F" };

                for (int i = 0, count = ColorList.Count; i < count; i++)
                {
                    brush = new SolidBrush(Color.Black);
                    sizeF = g.MeasureString(ColorList[i], font);
                    CurrentHeight += (int)(sizeF.Height) + 10;
                    g.DrawString(ColorList[i], font, brush, CenterOfLegend - LegendWidth / 2 + 50, CurrentHeight);

                    g.DrawRectangle(new Pen(Color.Black, 2.0f), CenterOfLegend - LegendWidth / 2 + 20 - 2, CurrentHeight - 2, RectWidth, RectHeight);

                    if (i == 0)
                    {
                        brush = new SolidBrush(Color.FromArgb(0x44, 0xff, 0x44));
                        g.FillRectangle(brush, CenterOfLegend - LegendWidth / 2 + 20 - 2, CurrentHeight - 2, RectWidth, RectHeight);
                    }
                    else if (i == 1)
                    {
                        brush = new SolidBrush(Color.FromArgb(0xff, 0x11, 0x11));
                        g.FillRectangle(brush, CenterOfLegend - LegendWidth / 2 + 20 - 2, CurrentHeight - 2, RectWidth, RectHeight);
                    }
                    else if (i == 2)
                    {
                        brush = new SolidBrush(Color.FromArgb(0xaa, 0xaa, 0xff));
                        g.FillRectangle(brush, CenterOfLegend - LegendWidth / 2 + 20 - 2, CurrentHeight - 2, RectWidth, RectHeight);
                    }
                }

                CurrentHeight += 40;

                string GoodLetterBad = "Good   A B C D E F   Bad";
                brush = new SolidBrush(Color.Black);
                sizeF = g.MeasureString(GoodLetterBad, font);
                g.DrawString(GoodLetterBad, font, brush, CenterOfLegend - (int)(sizeF.Width / 2), CurrentHeight);
            }
            catch (Exception ex)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotCreate_ImageError_, TaskRunnerServiceRes.Annotated, ex.Message + ex.InnerException != null ? " Inner: " + ex.InnerException.Message : "");
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotCreate_ImageError_", TaskRunnerServiceRes.Annotated, ex.Message + ex.InnerException != null ? " Inner: " + ex.InnerException.Message : "");
                return false;
            }


            return true;
        }
        private bool DrawLegendPolSourceSites(Graphics g, CoordMap coordMap, List<MapInfoPointModel> mapInfoPointModelPolSourceSiteList, List<TVItemModel> tvItemModelPolSourceSiteList)
        {
            int StartingHeight = 100;
            int CurrentHeight = 0;
            int CenterOfLegend = 0;
            Brush brush = new SolidBrush(Color.White);
            Pen pen = new Pen(Color.LightGreen, 6.0f);
            Font font = new Font("Arial", 16, FontStyle.Bold);

            int LegendHeight = 0;
            int LegendWidth = 200;
            // Calculating Total legend height
            string LegendText = TaskRunnerServiceRes.Legend;
            SizeF sizeF = g.MeasureString(LegendText, font);
            LegendHeight = LegendHeight + (int)(sizeF.Height) + 10;

            font = new Font("Arial", 12, FontStyle.Bold);
            string ApprovedText = "Some Text"; // just to measure the text height
            sizeF = g.MeasureString(ApprovedText, font);
            LegendHeight = LegendHeight + ((int)(sizeF.Height) + 5) * 13;

            g.DrawRectangle(pen, GoogleImageWidth * 2 - LegendWidth, StartingHeight, LegendWidth - 5, LegendHeight);
            g.FillRectangle(brush, GoogleImageWidth * 2 - LegendWidth, StartingHeight, LegendWidth - 5, LegendHeight);

            CenterOfLegend = (GoogleImageWidth * 2) - LegendWidth / 2;
            brush = new SolidBrush(Color.LightBlue);

            // draw Legend title
            font = new Font("Arial", 16, FontStyle.Bold);
            sizeF = g.MeasureString(LegendText, font);
            CurrentHeight = StartingHeight + 10;
            g.DrawString(LegendText, font, brush, CenterOfLegend - (sizeF.Width / 2), CurrentHeight);

            CurrentHeight += 20;
            font = new Font("Arial", 10, FontStyle.Bold);
            brush = new SolidBrush(Color.Blue);
            List<string> MajorGroupList = new List<string>()
            {
                TaskRunnerServiceRes.Agriculture, TaskRunnerServiceRes.Forested, TaskRunnerServiceRes.Industry,
                TaskRunnerServiceRes.Marine, TaskRunnerServiceRes.Recreation, TaskRunnerServiceRes.Urban, TaskRunnerServiceRes.Other
            };

            sizeF = g.MeasureString(MajorGroupList[0], font);
            int RectWidth = 24;
            int RectHeight = (int)(sizeF.Height) + 4;

            for (int i = 0, count = MajorGroupList.Count; i < count; i++)
            {
                brush = new SolidBrush(Color.Black);
                sizeF = g.MeasureString(MajorGroupList[i], font);
                CurrentHeight += (int)(sizeF.Height) + 10;
                g.DrawString(MajorGroupList[i], font, brush, CenterOfLegend - LegendWidth / 2 + 50, CurrentHeight);

                brush = new SolidBrush(Color.Green);
                int LatY = CurrentHeight + RectHeight / 2;
                int LngX = CenterOfLegend - LegendWidth / 2 + 20 + RectWidth / 2;
                switch (i)
                {
                    case 0: // Agriculture
                        {
                            DrawAgricultureIcon(g, new Pen(Color.Black, 1.0f), new SolidBrush(Color.Black), IconSize, LatY, LngX);
                        }
                        break;
                    case 1: // Forested
                        {
                            DrawForestedIcon(g, new Pen(Color.Black, 1.0f), new SolidBrush(Color.Black), IconSize, LatY, LngX);
                        }
                        break;
                    case 2: // Industry
                        {
                            DrawIndustryIcon(g, new Pen(Color.Black, 1.0f), new SolidBrush(Color.Black), IconSize, LatY, LngX);
                        }
                        break;
                    case 3: // Marine
                        {
                            DrawMarineIcon(g, new Pen(Color.Black, 1.0f), new SolidBrush(Color.Black), IconSize, LatY, LngX);
                        }
                        break;
                    case 4: // Recreation
                        {
                            DrawRecreationIcon(g, new Pen(Color.Black, 1.0f), new SolidBrush(Color.Black), IconSize, LatY, LngX);
                        }
                        break;
                    case 5: // Urban
                        {
                            DrawUrbanIcon(g, new Pen(Color.Black, 1.0f), new SolidBrush(Color.Black), IconSize, LatY, LngX);
                        }
                        break;
                    case 6: // Other
                        {
                            DrawOtherIcon(g, new Pen(Color.Black, 1.0f), new SolidBrush(Color.Black), IconSize, LatY, LngX);
                        }
                        break;
                    default:
                        break;
                }
            }

            CurrentHeight += 20;

            List<string> RiskList = new List<string>()
            {
               TaskRunnerServiceRes.HighRisk, TaskRunnerServiceRes.ModerateRisk, TaskRunnerServiceRes.LowRisk
            };

            for (int i = 0, count = RiskList.Count; i < count; i++)
            {
                sizeF = g.MeasureString(RiskList[i], font);
                CurrentHeight += (int)(sizeF.Height) + 10;

                if (i == 0)
                {
                    brush = new SolidBrush(Color.Red);
                }
                else if (i == 1)
                {
                    brush = new SolidBrush(Color.YellowGreen);
                }
                else if (i == 2)
                {
                    brush = new SolidBrush(Color.Green);
                }
                g.DrawString(RiskList[i], font, brush, CenterOfLegend - LegendWidth / 2 + 50, CurrentHeight);
            }

            return true;
        }
        private bool DrawMWQMSitesPoints(Graphics g, CoordMap coordMap, List<MapInfoPointModel> mapInfoPointModelMWQMSiteList, List<TVItemModel> tvItemModelMWQMSiteList)
        {
            string NotUsed = "";
            try
            {
                Brush brush = new SolidBrush(Color.LightGreen);
                Pen pen = new Pen(Color.LightGreen, 1.0f);
                Font font = new Font("Arial", 10, FontStyle.Regular);

                double TotalWidthLng = coordMap.NorthEast.Lng - coordMap.SouthWest.Lng;
                double TotalHeightLat = coordMap.NorthEast.Lat - coordMap.SouthWest.Lat;

                foreach (MapInfoPointModel mapInfoPointModel in mapInfoPointModelMWQMSiteList)
                {
                    TVItemModel tvItemModel = tvItemModelMWQMSiteList.Where(c => c.TVItemID == mapInfoPointModel.TVItemID).FirstOrDefault();

                    if (tvItemModel != null && tvItemModel.IsActive == true)
                    {
                        double LngX = ((mapInfoPointModel.Lng - coordMap.SouthWest.Lng) / TotalWidthLng) * GoogleImageWidth * 2.0D;
                        double LatY = ((GoogleImageHeight * 2) - GoogleLogoHeight) - ((TotalHeightLat - (coordMap.NorthEast.Lat - mapInfoPointModel.Lat)) / TotalHeightLat) * ((GoogleImageHeight * 2) - GoogleLogoHeight);

                        g.DrawRectangle(new Pen(Color.LightGreen, 1.0f), (int)LngX - 3, (int)LatY - 3, 6, 6);
                        g.FillRectangle(brush, (int)LngX - 3, (int)LatY - 3, 6, 6);

                        string TVText = tvItemModel.TVText;
                        while (true)
                        {
                            if (TVText.First().ToString() != "0")
                            {
                                break;
                            }
                            TVText = TVText.Substring(1);
                        }

                        g.DrawString(TVText, font, brush, (int)LngX, (int)LatY);

                    }
                }

            }
            catch (Exception ex)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotCreate_ImageError_, TaskRunnerServiceRes.Annotated, ex.Message + ex.InnerException != null ? " Inner: " + ex.InnerException.Message : "");
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotCreate_ImageError_", TaskRunnerServiceRes.Annotated, ex.Message + ex.InnerException != null ? " Inner: " + ex.InnerException.Message : "");
                return false;
            }
            return true;
        }
        private bool DrawNorthArrow(Graphics g, CoordMap coordMap)
        {
            string NotUsed = "";
            if (string.IsNullOrWhiteSpace(HideNorthArrow))
            {
                try
                {
                    string NorthText = "N";
                    Font font = new Font("Arial", 16, FontStyle.Bold);
                    Brush brush = new SolidBrush(Color.LightBlue);

                    SizeF sizeF = g.MeasureString(NorthText, font);

                    g.MeasureString(NorthText, font);
                    g.DrawString(NorthText, font, brush, GoogleImageWidth * 2 - 50 - (int)(sizeF.Width / 2), 30);

                    Pen pen = new Pen(Color.LightBlue, 6);
                    pen.StartCap = LineCap.NoAnchor;
                    pen.EndCap = LineCap.RoundAnchor;
                    g.DrawLine(pen, GoogleImageWidth * 2 - 50, 55, GoogleImageWidth * 2 - 50, 70);

                    pen = new Pen(Color.LightBlue, 6);
                    pen.StartCap = LineCap.ArrowAnchor;
                    pen.EndCap = LineCap.NoAnchor;
                    g.DrawLine(pen, GoogleImageWidth * 2 - 50, 15, GoogleImageWidth * 2 - 50, 30);

                }
                catch (Exception ex)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotCreate_ImageError_, TaskRunnerServiceRes.Annotated, ex.Message + ex.InnerException != null ? " Inner: " + ex.InnerException.Message : "");
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotCreate_ImageError_", TaskRunnerServiceRes.Annotated, ex.Message + ex.InnerException != null ? " Inner: " + ex.InnerException.Message : "");
                    return false;
                }
            }

            return true;
        }
        public bool DrawPolSourceSitesPoints(Graphics g, int GraphicWidth, int GraphicHeight, CoordMap coordMap, List<MapInfoPointModel> mapInfoPointModelPolSourceSiteList, List<TVItemModel> tvItemModelPolSourceSiteList)
        {
            float LabelHeightY = 0.0f;
            float LabelWidthX = 0.0f;
            float LabelHeightLat = 0.0f;
            float LabelWidthLng = 0.0f;

            Font font = new Font("Arial", 10, FontStyle.Regular);
            Brush brush = new SolidBrush(Color.LightGreen);

            SizeF sizeF = g.MeasureString("AA", font);
            LabelHeightY = sizeF.Height;
            LabelWidthX = sizeF.Width;

            // calculating LabelHeightLat and LabelWidthLng
            double TotalWidthLng = coordMap.NorthEast.Lng - coordMap.SouthWest.Lng;
            double TotalHeightLat = coordMap.NorthEast.Lat - coordMap.SouthWest.Lat;

            LabelHeightLat = (float)TotalHeightLat * (LabelHeightY / GraphicHeight);
            LabelWidthLng = (float)TotalWidthLng * (LabelWidthX / GraphicWidth);

            float StepSize = LabelHeightLat / 2.0f;

            if (tvItemModelPolSourceSiteList.Count > 0)
            {
                //List<LabelPosition> LabelPositionList = new List<LabelPosition>();

                //foreach (MapInfoPointModel mapInfoPointModel in mapInfoPointModelPolSourceSiteList)
                //{
                //    LabelPosition labelPosition = new LabelPosition()
                //    {
                //        TVItemID = mapInfoPointModel.TVItemID,
                //        SitePoint = new Coord() { Lat = (float)mapInfoPointModel.Lat, Lng = (float)mapInfoPointModel.Lng, Ordinal = 0 },
                //        LabelPoint = new Coord() { Lat = (float)mapInfoPointModel.Lat, Lng = (float)mapInfoPointModel.Lng, Ordinal = 0 },
                //        LabelNorthEast = new Coord() { Lat = (float)mapInfoPointModel.Lat + LabelHeightLat, Lng = (float)mapInfoPointModel.Lng + LabelWidthLng, Ordinal = 0 },
                //        LabelSouthWest = new Coord() { Lat = (float)mapInfoPointModel.Lat, Lng = (float)mapInfoPointModel.Lng, Ordinal = 0 },
                //        Position = PositionEnum.LeftBottom,
                //        Distance = 0.0f,
                //        Ordinal = LabelPositionList.Count(),
                //    };
                //    LabelPositionList.Add(labelPosition);
                //}

                //FillLabelPositionList(LabelPositionList, LabelHeightLat, LabelWidthLng, StepSize);

                List<PolSourceSiteModel> polSourceSiteModelList = _PolSourceSiteService.GetPolSourceSiteModelListWithSubsectorTVItemIDDB(tvItemModelPolSourceSiteList[0].ParentID);
                List<PolSourceObservationModel> polSourceObservationModelList = _PolSourceObservationService.GetPolSourceObservationModelListWithSubsectorTVItemIDDB(tvItemModelPolSourceSiteList[0].ParentID);
                List<PolSourceObservationIssueModel> polSourceObservationIssueModelList = _PolSourceObservationIssueService.GetPolSourceObservationIssueModelListWithSubsectorTVItemIDDB(tvItemModelPolSourceSiteList[0].ParentID);

                int count = 0;
                foreach (MapInfoPointModel mapInfoPointModel in mapInfoPointModelPolSourceSiteList)
                {
                    TVItemModel tvItemModel = tvItemModelPolSourceSiteList.Where(c => c.TVItemID == mapInfoPointModel.TVItemID).FirstOrDefault();

                    if (tvItemModel != null && tvItemModel.IsActive == true)
                    {
                        PolSourceSiteModel polSourceSiteModel = polSourceSiteModelList.Where(c => c.PolSourceSiteTVItemID == tvItemModel.TVItemID).FirstOrDefault();
                        if (polSourceSiteModel != null)
                        {
                            //LabelPosition labelPosition = LabelPositionList.Where(c => c.TVItemID == polSourceSiteModel.PolSourceSiteTVItemID).FirstOrDefault();
                            //if (labelPosition == null)
                            //{
                            //    continue;
                            //}
                            PolSourceObservationModel polSourceObservationModelLastest = polSourceObservationModelList.Where(c => c.PolSourceSiteID == polSourceSiteModel.PolSourceSiteID).OrderByDescending(c => c.ObservationDate_Local).FirstOrDefault();

                            if (polSourceObservationModelLastest != null)
                            {
                                count += 1;
                                double LngX = ((mapInfoPointModel.Lng - coordMap.SouthWest.Lng) / TotalWidthLng) * GoogleImageWidth * 2.0D;
                                double LatY = ((GoogleImageHeight * 2) - GoogleLogoHeight) - ((TotalHeightLat - (coordMap.NorthEast.Lat - mapInfoPointModel.Lat)) / TotalHeightLat) * ((GoogleImageHeight * 2) - GoogleLogoHeight);
                                //double LngXLabel = ((labelPosition.LabelPoint.Lng - coordMap.SouthWest.Lng) / TotalWidthLng) * GoogleImageWidth * 2.0D;
                                //double LatYLabel = ((GoogleImageHeight * 2) - GoogleLogoHeight) - ((TotalHeightLat - (coordMap.NorthEast.Lat - labelPosition.LabelPoint.Lat)) / TotalHeightLat) * ((GoogleImageHeight * 2) - GoogleLogoHeight);

                                //double LngXSW = ((labelPosition.LabelSouthWest.Lng - coordMap.SouthWest.Lng) / TotalWidthLng) * GoogleImageWidth * 2.0D;
                                //double LatYSW = ((GoogleImageHeight * 2) - GoogleLogoHeight) - ((TotalHeightLat - (coordMap.NorthEast.Lat - labelPosition.LabelSouthWest.Lat)) / TotalHeightLat) * ((GoogleImageHeight * 2) - GoogleLogoHeight);

                                //double LngXNE = ((labelPosition.LabelNorthEast.Lng - coordMap.SouthWest.Lng) / TotalWidthLng) * GoogleImageWidth * 2.0D;
                                //double LatYNE = ((GoogleImageHeight * 2) - GoogleLogoHeight) - ((TotalHeightLat - (coordMap.NorthEast.Lat - labelPosition.LabelNorthEast.Lat)) / TotalHeightLat) * ((GoogleImageHeight * 2) - GoogleLogoHeight);

                                //g.DrawLine(new Pen(Color.Green, 1.0f), (int)LngXLabel, (int)LatYLabel, (int)LngX, (int)LatY);
                                //g.DrawRectangle(new Pen(Color.Green, 1.0f), (int)LngXSW, (int)LatYSW, (int)LngXNE, (int)LatYNE);

                                List<PolSourceObservationIssueModel> polSourceObservationIssueModelListSelected = polSourceObservationIssueModelList.Where(c => c.PolSourceObservationID == polSourceObservationModelLastest.PolSourceObservationID).OrderBy(c => c.Ordinal).ToList();

                                //------------------------------------------------------------------------
                                // used this code to show map with moved labels
                                //------------------------------------------------------------------------
                                //if (polSourceObservationIssueModelListSelected.Count == 0)
                                //{
                                //    DrawOtherIcon(g, new Pen(Color.Black, 1.0f), new SolidBrush(Color.Black), IconSize, (int)((LatYNE + LatYSW)/2), (int)((LngXNE + LngXSW) / 2));
                                //}
                                //else
                                //{
                                //    if (polSourceObservationIssueModelListSelected[0].ObservationInfo.Contains(",10501,")) // Agriculture
                                //    {
                                //        if (polSourceObservationIssueModelListSelected[0].ObservationInfo.Contains(",93001,")
                                //            || polSourceObservationIssueModelListSelected[0].ObservationInfo.Contains(",92002,")) // High Risk
                                //        {
                                //            DrawAgricultureIcon(g, new Pen(Color.Red, 1.0f), new SolidBrush(Color.Red), IconSize, (int)((LatYNE + LatYSW) / 2), (int)((LngXNE + LngXSW) / 2));
                                //        }
                                //        else if (polSourceObservationIssueModelListSelected[0].ObservationInfo.Contains(",92001,")) // Moderate Risk
                                //        {
                                //            DrawAgricultureIcon(g, new Pen(Color.YellowGreen, 1.0f), new SolidBrush(Color.YellowGreen), IconSize, (int)((LatYNE + LatYSW) / 2), (int)((LngXNE + LngXSW) / 2));
                                //        }
                                //        else if (polSourceObservationIssueModelListSelected[0].ObservationInfo.Contains(",91001,")) // Low Risk
                                //        {
                                //            DrawAgricultureIcon(g, new Pen(Color.Green, 1.0f), new SolidBrush(Color.Green), IconSize, (int)((LatYNE + LatYSW) / 2), (int)((LngXNE + LngXSW) / 2));
                                //        }
                                //        else
                                //        {
                                //            DrawAgricultureIcon(g, new Pen(Color.Black, 1.0f), new SolidBrush(Color.Black), IconSize, (int)((LatYNE + LatYSW) / 2), (int)((LngXNE + LngXSW) / 2));
                                //        }
                                //    }
                                //    else if (polSourceObservationIssueModelListSelected[0].ObservationInfo.Contains(",10502,")) // Forested
                                //    {
                                //        if (polSourceObservationIssueModelListSelected[0].ObservationInfo.Contains(",93001,")
                                //            || polSourceObservationIssueModelListSelected[0].ObservationInfo.Contains(",92002,")) // High Risk
                                //        {
                                //            DrawForestedIcon(g, new Pen(Color.Red, 1.0f), new SolidBrush(Color.Red), IconSize, (int)((LatYNE + LatYSW) / 2), (int)((LngXNE + LngXSW) / 2));
                                //        }
                                //        else if (polSourceObservationIssueModelListSelected[0].ObservationInfo.Contains(",92001,")) // Moderate Risk
                                //        {
                                //            DrawForestedIcon(g, new Pen(Color.YellowGreen, 1.0f), new SolidBrush(Color.YellowGreen), IconSize, (int)((LatYNE + LatYSW) / 2), (int)((LngXNE + LngXSW) / 2));
                                //        }
                                //        else if (polSourceObservationIssueModelListSelected[0].ObservationInfo.Contains(",91001,")) // Low Risk
                                //        {
                                //            DrawForestedIcon(g, new Pen(Color.Green, 1.0f), new SolidBrush(Color.Green), IconSize, (int)((LatYNE + LatYSW) / 2), (int)((LngXNE + LngXSW) / 2));
                                //        }
                                //        else
                                //        {
                                //            DrawForestedIcon(g, new Pen(Color.Black, 1.0f), new SolidBrush(Color.Black), IconSize, (int)((LatYNE + LatYSW) / 2), (int)((LngXNE + LngXSW) / 2));
                                //        }
                                //    }
                                //    else if (polSourceObservationIssueModelListSelected[0].ObservationInfo.Contains(",10503,")) // Industry
                                //    {
                                //        if (polSourceObservationIssueModelListSelected[0].ObservationInfo.Contains(",93001,")
                                //            || polSourceObservationIssueModelListSelected[0].ObservationInfo.Contains(",92002,")) // High Risk
                                //        {
                                //            DrawIndustryIcon(g, new Pen(Color.Red, 1.0f), new SolidBrush(Color.Red), IconSize, (int)((LatYNE + LatYSW) / 2), (int)((LngXNE + LngXSW) / 2));
                                //        }
                                //        else if (polSourceObservationIssueModelListSelected[0].ObservationInfo.Contains(",92001,")) // Moderate Risk
                                //        {
                                //            DrawIndustryIcon(g, new Pen(Color.YellowGreen, 1.0f), new SolidBrush(Color.YellowGreen), IconSize, (int)((LatYNE + LatYSW) / 2), (int)((LngXNE + LngXSW) / 2));
                                //        }
                                //        else if (polSourceObservationIssueModelListSelected[0].ObservationInfo.Contains(",91001,")) // Low Risk
                                //        {
                                //            DrawIndustryIcon(g, new Pen(Color.Green, 1.0f), new SolidBrush(Color.Green), IconSize, (int)((LatYNE + LatYSW) / 2), (int)((LngXNE + LngXSW) / 2));
                                //        }
                                //        else
                                //        {
                                //            DrawIndustryIcon(g, new Pen(Color.Black, 1.0f), new SolidBrush(Color.Black), IconSize, (int)((LatYNE + LatYSW) / 2), (int)((LngXNE + LngXSW) / 2));
                                //        }
                                //    }
                                //    else if (polSourceObservationIssueModelListSelected[0].ObservationInfo.Contains(",10504,")) // Marine
                                //    {
                                //        if (polSourceObservationIssueModelListSelected[0].ObservationInfo.Contains(",93001,")
                                //            || polSourceObservationIssueModelListSelected[0].ObservationInfo.Contains(",92002,")) // High Risk
                                //        {
                                //            DrawMarineIcon(g, new Pen(Color.Red, 1.0f), new SolidBrush(Color.Red), IconSize, (int)((LatYNE + LatYSW)/2), (int)((LngXNE + LngXSW) / 2));
                                //        }
                                //        else if (polSourceObservationIssueModelListSelected[0].ObservationInfo.Contains(",92001,")) // Moderate Risk
                                //        {
                                //            DrawMarineIcon(g, new Pen(Color.YellowGreen, 1.0f), new SolidBrush(Color.YellowGreen), IconSize, (int)((LatYNE + LatYSW)/2), (int)((LngXNE + LngXSW) / 2));
                                //        }
                                //        else if (polSourceObservationIssueModelListSelected[0].ObservationInfo.Contains(",91001,")) // Low Risk
                                //        {
                                //            DrawMarineIcon(g, new Pen(Color.Green, 1.0f), new SolidBrush(Color.Green), IconSize, (int)((LatYNE + LatYSW)/2), (int)((LngXNE + LngXSW) / 2));
                                //        }
                                //        else
                                //        {
                                //            DrawMarineIcon(g, new Pen(Color.Black, 1.0f), new SolidBrush(Color.Black), IconSize, (int)((LatYNE + LatYSW)/2), (int)((LngXNE + LngXSW) / 2));
                                //        }
                                //    }
                                //    else if (polSourceObservationIssueModelListSelected[0].ObservationInfo.Contains(",10505,")) // Recreation
                                //    {
                                //        if (polSourceObservationIssueModelListSelected[0].ObservationInfo.Contains(",93001,")
                                //            || polSourceObservationIssueModelListSelected[0].ObservationInfo.Contains(",92002,")) // High Risk
                                //        {
                                //            DrawRecreationIcon(g, new Pen(Color.Red, 1.0f), new SolidBrush(Color.Red), IconSize, (int)((LatYNE + LatYSW)/2), (int)((LngXNE + LngXSW) / 2));
                                //        }
                                //        else if (polSourceObservationIssueModelListSelected[0].ObservationInfo.Contains(",92001,")) // Moderate Risk
                                //        {
                                //            DrawRecreationIcon(g, new Pen(Color.YellowGreen, 1.0f), new SolidBrush(Color.YellowGreen), IconSize, (int)((LatYNE + LatYSW)/2), (int)((LngXNE + LngXSW) / 2));
                                //        }
                                //        else if (polSourceObservationIssueModelListSelected[0].ObservationInfo.Contains(",91001,")) // Low Risk
                                //        {
                                //            DrawRecreationIcon(g, new Pen(Color.Green, 1.0f), new SolidBrush(Color.Green), IconSize, (int)((LatYNE + LatYSW)/2), (int)((LngXNE + LngXSW) / 2));
                                //        }
                                //        else
                                //        {
                                //            DrawRecreationIcon(g, new Pen(Color.Black, 1.0f), new SolidBrush(Color.Black), IconSize, (int)((LatYNE + LatYSW)/2), (int)((LngXNE + LngXSW) / 2));
                                //        }
                                //    }
                                //    else if (polSourceObservationIssueModelListSelected[0].ObservationInfo.Contains(",10506,")) // Urban
                                //    {
                                //        if (polSourceObservationIssueModelListSelected[0].ObservationInfo.Contains(",93001,")
                                //            || polSourceObservationIssueModelListSelected[0].ObservationInfo.Contains(",92002,")) // High Risk
                                //        {
                                //            DrawUrbanIcon(g, new Pen(Color.Red, 1.0f), new SolidBrush(Color.Red), IconSize, (int)((LatYNE + LatYSW)/2), (int)((LngXNE + LngXSW) / 2));
                                //        }
                                //        else if (polSourceObservationIssueModelListSelected[0].ObservationInfo.Contains(",92001,")) // Moderate Risk
                                //        {
                                //            DrawUrbanIcon(g, new Pen(Color.YellowGreen, 1.0f), new SolidBrush(Color.YellowGreen), IconSize, (int)((LatYNE + LatYSW)/2), (int)((LngXNE + LngXSW) / 2));
                                //        }
                                //        else if (polSourceObservationIssueModelListSelected[0].ObservationInfo.Contains(",91001,")) // Low Risk
                                //        {
                                //            DrawUrbanIcon(g, new Pen(Color.Green, 1.0f), new SolidBrush(Color.Green), IconSize, (int)((LatYNE + LatYSW)/2), (int)((LngXNE + LngXSW) / 2));
                                //        }
                                //        else
                                //        {
                                //            DrawUrbanIcon(g, new Pen(Color.Black, 1.0f), new SolidBrush(Color.Black), IconSize, (int)((LatYNE + LatYSW)/2), (int)((LngXNE + LngXSW) / 2));
                                //        }
                                //    }
                                //    else
                                //    {
                                //        DrawOtherIcon(g, new Pen(Color.Black, 1.0f), new SolidBrush(Color.Black), IconSize, (int)((LatYNE + LatYSW)/2), (int)((LngXNE + LngXSW) / 2));
                                //    }
                                //}

                                //------------------------------------------------------------------------
                                // used this code to show map No moved labels
                                //------------------------------------------------------------------------
                                if (polSourceObservationIssueModelListSelected.Count == 0)
                                {
                                    DrawOtherIcon(g, new Pen(Color.Black, 1.0f), new SolidBrush(Color.Black), IconSize, (int)LatY, (int)LngX);
                                }
                                else
                                {
                                    if (polSourceObservationIssueModelListSelected[0].ObservationInfo.Contains(",10501,")) // Agriculture
                                    {
                                        if (polSourceObservationIssueModelListSelected[0].ObservationInfo.Contains(",93001,")
                                            || polSourceObservationIssueModelListSelected[0].ObservationInfo.Contains(",92002,")) // High Risk
                                        {
                                            DrawAgricultureIcon(g, new Pen(Color.Red, 1.0f), new SolidBrush(Color.Red), IconSize, (int)LatY, (int)LngX);
                                        }
                                        else if (polSourceObservationIssueModelListSelected[0].ObservationInfo.Contains(",92001,")) // Moderate Risk
                                        {
                                            DrawAgricultureIcon(g, new Pen(Color.YellowGreen, 1.0f), new SolidBrush(Color.YellowGreen), IconSize, (int)LatY, (int)LngX);
                                        }
                                        else if (polSourceObservationIssueModelListSelected[0].ObservationInfo.Contains(",91001,")) // Low Risk
                                        {
                                            DrawAgricultureIcon(g, new Pen(Color.Green, 1.0f), new SolidBrush(Color.Green), IconSize, (int)LatY, (int)LngX);
                                        }
                                        else
                                        {
                                            DrawAgricultureIcon(g, new Pen(Color.Black, 1.0f), new SolidBrush(Color.Black), IconSize, (int)LatY, (int)LngX);
                                        }
                                    }
                                    else if (polSourceObservationIssueModelListSelected[0].ObservationInfo.Contains(",10502,")) // Forested
                                    {
                                        if (polSourceObservationIssueModelListSelected[0].ObservationInfo.Contains(",93001,")
                                            || polSourceObservationIssueModelListSelected[0].ObservationInfo.Contains(",92002,")) // High Risk
                                        {
                                            DrawForestedIcon(g, new Pen(Color.Red, 1.0f), new SolidBrush(Color.Red), IconSize, (int)LatY, (int)LngX);
                                        }
                                        else if (polSourceObservationIssueModelListSelected[0].ObservationInfo.Contains(",92001,")) // Moderate Risk
                                        {
                                            DrawForestedIcon(g, new Pen(Color.YellowGreen, 1.0f), new SolidBrush(Color.YellowGreen), IconSize, (int)LatY, (int)LngX);
                                        }
                                        else if (polSourceObservationIssueModelListSelected[0].ObservationInfo.Contains(",91001,")) // Low Risk
                                        {
                                            DrawForestedIcon(g, new Pen(Color.Green, 1.0f), new SolidBrush(Color.Green), IconSize, (int)LatY, (int)LngX);
                                        }
                                        else
                                        {
                                            DrawForestedIcon(g, new Pen(Color.Black, 1.0f), new SolidBrush(Color.Black), IconSize, (int)LatY, (int)LngX);
                                        }
                                    }
                                    else if (polSourceObservationIssueModelListSelected[0].ObservationInfo.Contains(",10503,")) // Industry
                                    {
                                        if (polSourceObservationIssueModelListSelected[0].ObservationInfo.Contains(",93001,")
                                            || polSourceObservationIssueModelListSelected[0].ObservationInfo.Contains(",92002,")) // High Risk
                                        {
                                            DrawIndustryIcon(g, new Pen(Color.Red, 1.0f), new SolidBrush(Color.Red), IconSize, (int)LatY, (int)LngX);
                                        }
                                        else if (polSourceObservationIssueModelListSelected[0].ObservationInfo.Contains(",92001,")) // Moderate Risk
                                        {
                                            DrawIndustryIcon(g, new Pen(Color.YellowGreen, 1.0f), new SolidBrush(Color.YellowGreen), IconSize, (int)LatY, (int)LngX);
                                        }
                                        else if (polSourceObservationIssueModelListSelected[0].ObservationInfo.Contains(",91001,")) // Low Risk
                                        {
                                            DrawIndustryIcon(g, new Pen(Color.Green, 1.0f), new SolidBrush(Color.Green), IconSize, (int)LatY, (int)LngX);
                                        }
                                        else
                                        {
                                            DrawIndustryIcon(g, new Pen(Color.Black, 1.0f), new SolidBrush(Color.Black), IconSize, (int)LatY, (int)LngX);
                                        }
                                    }
                                    else if (polSourceObservationIssueModelListSelected[0].ObservationInfo.Contains(",10504,")) // Marine
                                    {
                                        if (polSourceObservationIssueModelListSelected[0].ObservationInfo.Contains(",93001,")
                                            || polSourceObservationIssueModelListSelected[0].ObservationInfo.Contains(",92002,")) // High Risk
                                        {
                                            DrawMarineIcon(g, new Pen(Color.Red, 1.0f), new SolidBrush(Color.Red), IconSize, (int)LatY, (int)LngX);
                                        }
                                        else if (polSourceObservationIssueModelListSelected[0].ObservationInfo.Contains(",92001,")) // Moderate Risk
                                        {
                                            DrawMarineIcon(g, new Pen(Color.YellowGreen, 1.0f), new SolidBrush(Color.YellowGreen), IconSize, (int)LatY, (int)LngX);
                                        }
                                        else if (polSourceObservationIssueModelListSelected[0].ObservationInfo.Contains(",91001,")) // Low Risk
                                        {
                                            DrawMarineIcon(g, new Pen(Color.Green, 1.0f), new SolidBrush(Color.Green), IconSize, (int)LatY, (int)LngX);
                                        }
                                        else
                                        {
                                            DrawMarineIcon(g, new Pen(Color.Black, 1.0f), new SolidBrush(Color.Black), IconSize, (int)LatY, (int)LngX);
                                        }
                                    }
                                    else if (polSourceObservationIssueModelListSelected[0].ObservationInfo.Contains(",10505,")) // Recreation
                                    {
                                        if (polSourceObservationIssueModelListSelected[0].ObservationInfo.Contains(",93001,")
                                            || polSourceObservationIssueModelListSelected[0].ObservationInfo.Contains(",92002,")) // High Risk
                                        {
                                            DrawRecreationIcon(g, new Pen(Color.Red, 1.0f), new SolidBrush(Color.Red), IconSize, (int)LatY, (int)LngX);
                                        }
                                        else if (polSourceObservationIssueModelListSelected[0].ObservationInfo.Contains(",92001,")) // Moderate Risk
                                        {
                                            DrawRecreationIcon(g, new Pen(Color.YellowGreen, 1.0f), new SolidBrush(Color.YellowGreen), IconSize, (int)LatY, (int)LngX);
                                        }
                                        else if (polSourceObservationIssueModelListSelected[0].ObservationInfo.Contains(",91001,")) // Low Risk
                                        {
                                            DrawRecreationIcon(g, new Pen(Color.Green, 1.0f), new SolidBrush(Color.Green), IconSize, (int)LatY, (int)LngX);
                                        }
                                        else
                                        {
                                            DrawRecreationIcon(g, new Pen(Color.Black, 1.0f), new SolidBrush(Color.Black), IconSize, (int)LatY, (int)LngX);
                                        }
                                    }
                                    else if (polSourceObservationIssueModelListSelected[0].ObservationInfo.Contains(",10506,")) // Urban
                                    {
                                        if (polSourceObservationIssueModelListSelected[0].ObservationInfo.Contains(",93001,")
                                            || polSourceObservationIssueModelListSelected[0].ObservationInfo.Contains(",92002,")) // High Risk
                                        {
                                            DrawUrbanIcon(g, new Pen(Color.Red, 1.0f), new SolidBrush(Color.Red), IconSize, (int)LatY, (int)LngX);
                                        }
                                        else if (polSourceObservationIssueModelListSelected[0].ObservationInfo.Contains(",92001,")) // Moderate Risk
                                        {
                                            DrawUrbanIcon(g, new Pen(Color.YellowGreen, 1.0f), new SolidBrush(Color.YellowGreen), IconSize, (int)LatY, (int)LngX);
                                        }
                                        else if (polSourceObservationIssueModelListSelected[0].ObservationInfo.Contains(",91001,")) // Low Risk
                                        {
                                            DrawUrbanIcon(g, new Pen(Color.Green, 1.0f), new SolidBrush(Color.Green), IconSize, (int)LatY, (int)LngX);
                                        }
                                        else
                                        {
                                            DrawUrbanIcon(g, new Pen(Color.Black, 1.0f), new SolidBrush(Color.Black), IconSize, (int)LatY, (int)LngX);
                                        }
                                    }
                                    else
                                    {
                                        DrawOtherIcon(g, new Pen(Color.Black, 1.0f), new SolidBrush(Color.Black), IconSize, (int)LatY, (int)LngX);
                                    }
                                }
                                g.DrawString(count.ToString(), font, brush, new Point((int)(LngX + 2), (int)(LatY + 5)));
                            }
                        }
                    }
                }
            }

            return true;
        }
        private bool DrawSubsectorMWQMSites(CoordMap coordMap, TVItemModel tvItemModelSubsector, List<MapInfoPointModel> mapInfoPointModelSubsectorList, List<MapInfoPointModel> mapInfoPointModelMWQMSiteList, List<TVItemModel> tvItemModelMWQMSiteList)
        {
            string NotUsed = "";

            if (!GetInset(coordMap.NorthEast, coordMap.SouthWest))
            {
                return false;
            }

            try
            {
                using (Bitmap targetAll = new Bitmap(DirName + FileNameFull))
                {
                    using (Graphics g = Graphics.FromImage(targetAll))
                    {
                        using (Bitmap targetImg = new Bitmap(DirName + FileNameInsetFinal))
                        {
                            g.DrawImage(targetImg, new Point(0, 0));
                        }

                        if (!DrawImageBorder(g))
                        {
                            return false;
                        }

                        if (!DrawSubsectorPolygon(g, coordMap, mapInfoPointModelSubsectorList))
                        {
                            return false;
                        }

                        if (!DrawTitles(g, tvItemModelSubsector.TVText, TaskRunnerServiceRes.MarineWaterQualityMonitoringSites))
                        {
                            return false;
                        }

                        //if (!DrawLegendMWQMSites(g, coordMap, mapInfoPointModelMWQMSiteList, tvItemModelMWQMSiteList))
                        //{
                        //    return false;
                        //}

                        if (!DrawNorthArrow(g, coordMap))
                        {
                            return false;
                        }

                        if (!DrawHorizontalScale(g, coordMap))
                        {
                            return false;
                        }

                        if (!DrawVerticalScale(g, coordMap))
                        {
                            return false;
                        }

                        if (!DrawMWQMSitesPoints(g, coordMap, mapInfoPointModelMWQMSiteList, tvItemModelMWQMSiteList))
                        {
                            return false;
                        }

                    }

                    targetAll.Save(DirName + FileNameFull.Replace(".png", "Annotated.png"), ImageFormat.Png);
                }
            }
            catch (Exception ex)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotCreate_ImageError_, TaskRunnerServiceRes.Annotated, ex.Message + ex.InnerException != null ? " Inner: " + ex.InnerException.Message : "");
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotCreate_ImageError_", TaskRunnerServiceRes.Annotated, ex.Message + ex.InnerException != null ? " Inner: " + ex.InnerException.Message : "");
                return false;
            }

            return true;
        }
        private bool DrawSubsectorPolygon(Graphics g, CoordMap coordMap, List<MapInfoPointModel> mapInfoPointModelSubsectorList)
        {
            string NotUsed = "";
            try
            {
                List<Point> polygonPointList = new List<Point>();

                double TotalWidthLng = coordMap.NorthEast.Lng - coordMap.SouthWest.Lng;
                double TotalHeightLat = coordMap.NorthEast.Lat - coordMap.SouthWest.Lat;

                foreach (MapInfoPointModel mapInfoPointModel in mapInfoPointModelSubsectorList)
                {
                    double LngX = ((mapInfoPointModel.Lng - coordMap.SouthWest.Lng) / TotalWidthLng) * GoogleImageWidth * 2.0D;
                    double LatY = ((GoogleImageHeight * 2) - GoogleLogoHeight) - ((TotalHeightLat - (coordMap.NorthEast.Lat - mapInfoPointModel.Lat)) / TotalHeightLat) * ((GoogleImageHeight * 2) - GoogleLogoHeight);
                    polygonPointList.Add(new Point() { X = (int)LngX, Y = (int)LatY });
                }

                g.DrawPolygon(new Pen(Color.Orange, 2.0f), polygonPointList.ToArray());
            }
            catch (Exception ex)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotCreate_ImageError_, TaskRunnerServiceRes.Annotated, ex.Message + ex.InnerException != null ? " Inner: " + ex.InnerException.Message : "");
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotCreate_ImageError_", TaskRunnerServiceRes.Annotated, ex.Message + ex.InnerException != null ? " Inner: " + ex.InnerException.Message : "");
                return false;
            }

            return true;
        }
        private bool DrawSubsectorPolSourceSites(CoordMap coordMap, TVItemModel tvItemModelSubsector, List<MapInfoPointModel> mapInfoPointModelSubsectorList, List<MapInfoPointModel> mapInfoPointModelPolSourceSiteList, List<TVItemModel> tvItemModelPolSourceSiteList)
        {
            if (!GetInset(coordMap.NorthEast, coordMap.SouthWest))
            {
                return false;
            }

            using (Bitmap targetAll = new Bitmap(DirName + FileNameFull))
            {
                int GraphicWidth = targetAll.Width;
                int GraphicHeight = targetAll.Height;

                using (Graphics g = Graphics.FromImage(targetAll))
                {
                    using (Bitmap targetImg = new Bitmap(DirName + FileNameInsetFinal))
                    {
                        g.DrawImage(targetImg, new Point(0, 0));
                    }

                    if (!DrawImageBorder(g))
                    {
                        return false;
                    }

                    if (!DrawSubsectorPolygon(g, coordMap, mapInfoPointModelSubsectorList))
                    {
                        return false;
                    }

                    if (!DrawTitles(g, tvItemModelSubsector.TVText, TaskRunnerServiceRes.PollutionSourceSites))
                    {
                        return false;
                    }

                    if (!DrawLegendPolSourceSites(g, coordMap, mapInfoPointModelPolSourceSiteList, tvItemModelPolSourceSiteList))
                    {
                        return false;
                    }

                    if (!DrawNorthArrow(g, coordMap))
                    {
                        return false;
                    }

                    if (!DrawHorizontalScale(g, coordMap))
                    {
                        return false;
                    }

                    if (!DrawVerticalScale(g, coordMap))
                    {
                        return false;
                    }

                    if (!DrawPolSourceSitesPoints(g, GraphicWidth, GraphicHeight, coordMap, mapInfoPointModelPolSourceSiteList, tvItemModelPolSourceSiteList))
                    {
                        return false;
                    }

                }

                targetAll.Save(DirName + FileNameFullAnnotated, ImageFormat.Png);
            }

            return true;
        }
        private bool DrawTitles(Graphics g, string FirstTitle, string SecondTitle)
        {
            string NotUsed = "";
            try
            {
                float FirstTitleHeight = 0;
                Font font = new Font("Arial", 20, FontStyle.Bold);
                Brush brush = new SolidBrush(Color.LightBlue);

                SizeF sizeFInit = g.MeasureString(FirstTitle, font);
                SizeF sizeF = g.MeasureString(FirstTitle, font);
                if (string.IsNullOrWhiteSpace(HideSubsectorName))
                {
                    // First Title
                    while (true)
                    {
                        sizeF = g.MeasureString(FirstTitle, font);
                        if (sizeF.Width > (GoogleImageWidth * 2 - 200 - 200 - 100)) // 200 is the Inset and Legend width
                        {
                            FirstTitle = FirstTitle.Substring(0, FirstTitle.Length - 2);
                        }
                        else
                        {
                            break;
                        }
                    }
                    if (sizeFInit.Width != sizeF.Width)
                    {
                        FirstTitle = FirstTitle + "...";
                    }
                    g.DrawString(FirstTitle, font, brush, new Point(GoogleImageWidth - (int)(sizeF.Width / 2), 3));

                    FirstTitleHeight = sizeF.Height;
                }

                // Second Title
                font = new Font("Arial", 16, FontStyle.Bold);
                brush = new SolidBrush(Color.LightBlue);
                sizeFInit = g.MeasureString(SecondTitle, font);
                sizeF = g.MeasureString(SecondTitle, font);
                while (true)
                {
                    sizeF = g.MeasureString(SecondTitle, font);
                    if (sizeF.Width > (GoogleImageWidth * 2 - 200 - 200 - 100)) // 200 is the Inset and Legend width
                    {
                        SecondTitle = SecondTitle.Substring(0, SecondTitle.Length - 2);
                    }
                    else
                    {
                        break;
                    }
                }
                if (sizeFInit.Width != sizeF.Width)
                {
                    SecondTitle = SecondTitle + "...";
                }
                g.DrawString(SecondTitle, font, brush, new Point(GoogleImageWidth - (int)(sizeF.Width / 2), (int)((3.0f * 2.0f) + FirstTitleHeight)));
            }
            catch (Exception ex)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotCreate_ImageError_, TaskRunnerServiceRes.Annotated, ex.Message + ex.InnerException != null ? " Inner: " + ex.InnerException.Message : "");
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotCreate_ImageError_", TaskRunnerServiceRes.Annotated, ex.Message + ex.InnerException != null ? " Inner: " + ex.InnerException.Message : "");
                return false;
            }

            return true;
        }
        private bool DrawVerticalScale(Graphics g, CoordMap coordMap)
        {
            string NotUsed = "";
            if (string.IsNullOrWhiteSpace(HideVerticalScale))
            {
                try
                {
                    int MinLatY = (GoogleImageHeight * 2 - GoogleLogoHeight) - 60;
                    int MaxLatY = ((GoogleImageHeight * 2 - GoogleLogoHeight) * 4 / 5) - 60;

                    g.DrawLine(new Pen(Color.LightBlue, 2), 40, MinLatY, 40, MaxLatY);

                    double MinLat = coordMap.NorthEast.Lat - (coordMap.NorthEast.Lat - coordMap.SouthWest.Lat) / 5;
                    double distLat = _MapInfoService.CalculateDistance(MinLat * _MapInfoService.d2r, coordMap.NorthEast.Lng * _MapInfoService.d2r, coordMap.NorthEast.Lat * _MapInfoService.d2r, coordMap.NorthEast.Lng * _MapInfoService.d2r, _MapInfoService.R) / 1000;

                    string distText = distLat.ToString("F2") + " km";
                    Font font = new Font("Arial", 10, FontStyle.Regular);
                    Brush brush = new SolidBrush(Color.LightBlue);

                    SizeF sizeF = g.MeasureString(distText, font);

                    g.DrawString(distText, font, brush, 45, ((GoogleImageHeight * 2 - GoogleLogoHeight) * 4 / 5) - 60 - sizeF.Height);

                    for (int i = 0; i < 100; i++)
                    {
                        if ((double)i > distLat)
                        {
                            g.DrawLine(new Pen(Color.LightBlue, 1), 40 - 1, MaxLatY, 40 - 1 - 10, MaxLatY);
                            break;
                        }
                        g.DrawLine(new Pen(Color.LightBlue, 1), 40 - 1, MinLatY + (int)(i / distLat * (MaxLatY - MinLatY)), 40 - 1 - 10, MinLatY + (int)(i / distLat * (MaxLatY - MinLatY)));

                        distText = i.ToString();
                        font = new Font("Arial", 10, FontStyle.Regular);
                        brush = new SolidBrush(Color.LightBlue);

                        sizeF = g.MeasureString(distText, font);

                        g.DrawString(distText, font, brush, 45, MinLatY + (int)(i / distLat * (MaxLatY - MinLatY)) - (sizeF.Height / 2));
                    }
                }
                catch (Exception ex)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotCreate_ImageError_, TaskRunnerServiceRes.Annotated, ex.Message + ex.InnerException != null ? " Inner: " + ex.InnerException.Message : "");
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotCreate_ImageError_", TaskRunnerServiceRes.Annotated, ex.Message + ex.InnerException != null ? " Inner: " + ex.InnerException.Message : "");
                    return false;
                }
            }

            return true;
        }
        private void FillLabelPositionList(List<LabelPosition> LabelPositionList, float LabelHeight, float LabelWidth, float StepSize)
        {
            if (_MapInfoService == null)
            {
                _MapInfoService = new MapInfoService(LanguageEnum.en, null);
            }
            if (LabelPositionList.Count > 0)
            {
                PointF AveragePoint = new PointF(LabelPositionList.Average(c => c.SitePoint.Lng), LabelPositionList.Average(c => c.SitePoint.Lat));

                foreach (LabelPosition labelPosition in LabelPositionList)
                {
                    labelPosition.Distance = (float)Math.Sqrt((labelPosition.SitePoint.Lng - AveragePoint.X) * (labelPosition.SitePoint.Lng - AveragePoint.X) + (labelPosition.SitePoint.Lat - AveragePoint.Y) * (labelPosition.SitePoint.Lat - AveragePoint.Y));

                    if ((labelPosition.SitePoint.Lng - AveragePoint.X) >= 0 && (labelPosition.SitePoint.Lat - AveragePoint.Y) >= 0) // first quartier
                    {
                        labelPosition.LabelSouthWest = new Coord() { Lat = labelPosition.SitePoint.Lat + (LabelHeight / 2), Lng = labelPosition.SitePoint.Lng + (LabelWidth / 2), Ordinal = 0 };
                        labelPosition.LabelNorthEast = new Coord() { Lat = labelPosition.SitePoint.Lat + (LabelHeight * 3 / 2), Lng = labelPosition.SitePoint.Lng + (LabelWidth * 3 / 2), Ordinal = 0 };
                        labelPosition.Position = PositionEnum.LeftBottom;
                        labelPosition.LabelPoint = new Coord() { Lat = labelPosition.LabelSouthWest.Lat, Lng = labelPosition.LabelSouthWest.Lng, Ordinal = 0 };
                    }
                    else if ((labelPosition.SitePoint.Lng - AveragePoint.X) > 0 && (labelPosition.SitePoint.Lat - AveragePoint.Y) < 0) // second quartier
                    {
                        labelPosition.LabelSouthWest = new Coord() { Lat = labelPosition.SitePoint.Lat - (LabelHeight * 3 / 2), Lng = labelPosition.SitePoint.Lng + (LabelWidth / 2), Ordinal = 0 };
                        labelPosition.LabelNorthEast = new Coord() { Lat = labelPosition.SitePoint.Lat - (LabelHeight / 2), Lng = labelPosition.SitePoint.Lng + (LabelWidth * 3 / 2), Ordinal = 0 };
                        labelPosition.Position = PositionEnum.LeftTop;
                        labelPosition.LabelPoint = new Coord() { Lat = labelPosition.LabelNorthEast.Lat, Lng = labelPosition.LabelSouthWest.Lng, Ordinal = 0 };
                    }
                    else if ((labelPosition.SitePoint.Lng - AveragePoint.X) < 0 && (labelPosition.SitePoint.Lat - AveragePoint.Y) < 0) // third quartier
                    {
                        labelPosition.LabelSouthWest = new Coord() { Lat = labelPosition.SitePoint.Lat - (LabelHeight * 3 / 2), Lng = labelPosition.SitePoint.Lng - (LabelWidth * 3 / 2), Ordinal = 0 };
                        labelPosition.LabelNorthEast = new Coord() { Lat = labelPosition.SitePoint.Lat - (LabelHeight / 2), Lng = labelPosition.SitePoint.Lng - (LabelWidth / 2), Ordinal = 0 };
                        labelPosition.Position = PositionEnum.RightTop;
                        labelPosition.LabelPoint = new Coord() { Lat = labelPosition.LabelNorthEast.Lat, Lng = labelPosition.LabelNorthEast.Lng, Ordinal = 0 };
                    }
                    else // forth quartier
                    {
                        labelPosition.LabelSouthWest = new Coord() { Lat = labelPosition.SitePoint.Lat + (LabelHeight / 2), Lng = labelPosition.SitePoint.Lng - (LabelWidth * 3 / 2), Ordinal = 0 };
                        labelPosition.LabelNorthEast = new Coord() { Lat = labelPosition.SitePoint.Lat + (LabelHeight * 3 / 2), Lng = labelPosition.SitePoint.Lng - (LabelWidth / 2), Ordinal = 0 };
                        labelPosition.Position = PositionEnum.RightBottom;
                        labelPosition.LabelPoint = new Coord() { Lat = labelPosition.LabelSouthWest.Lat, Lng = labelPosition.LabelNorthEast.Lng, Ordinal = 0 };
                    }
                }
                foreach (LabelPosition labelPosition in LabelPositionList.OrderBy(c => c.Distance))
                {
                    bool HidingPoint = true;
                    while (HidingPoint)
                    {
                        List<Coord> coordList = new List<Coord>()
                        {
                            new Coord() { Lat = labelPosition.LabelSouthWest.Lat, Lng = labelPosition.LabelSouthWest.Lng, Ordinal = 0 },
                            new Coord() { Lat = labelPosition.LabelSouthWest.Lat, Lng = labelPosition.LabelNorthEast.Lng, Ordinal = 0 },
                            new Coord() { Lat = labelPosition.LabelNorthEast.Lat, Lng = labelPosition.LabelNorthEast.Lng, Ordinal = 0 },
                            new Coord() { Lat = labelPosition.LabelNorthEast.Lat, Lng = labelPosition.LabelSouthWest.Lng, Ordinal = 0 },
                        };

                        bool PleaseRedo = false;
                        foreach (LabelPosition labelPosition2 in LabelPositionList.Where(c => c.Ordinal != labelPosition.Ordinal))
                        {
                            Coord coord = new Coord()
                            {
                                Lat = labelPosition2.LabelPoint.Lat,
                                Lng = labelPosition2.LabelPoint.Lng,
                                Ordinal = 0,
                            };
                            if (_MapInfoService.CoordInPolygon(coordList, coord))
                            {
                                float XNew = StepSize;
                                float YNew = StepSize;
                                float dist = (float)Math.Sqrt((AveragePoint.Y - labelPosition.SitePoint.Lat) * (AveragePoint.Y - labelPosition.SitePoint.Lat) + (AveragePoint.X - labelPosition.SitePoint.Lng) * (AveragePoint.X - labelPosition.SitePoint.Lng));
                                float factor = dist / StepSize;
                                float deltX = Math.Abs((AveragePoint.X - labelPosition.LabelPoint.Lng) / factor);
                                float deltY = Math.Abs((AveragePoint.Y - labelPosition.LabelPoint.Lat) / factor);
                                switch (labelPosition.Position)
                                {
                                    case PositionEnum.Error:
                                        break;
                                    case PositionEnum.LeftBottom:
                                        {
                                            XNew = labelPosition.LabelPoint.Lng + deltX;
                                            YNew = labelPosition.LabelPoint.Lat - deltY;
                                            labelPosition.LabelSouthWest = new Coord() { Lat = YNew, Lng = XNew, Ordinal = 0 };
                                            labelPosition.LabelNorthEast = new Coord() { Lat = YNew - LabelHeight, Lng = XNew + LabelWidth, Ordinal = 0 };
                                        }
                                        break;
                                    case PositionEnum.RightBottom:
                                        {
                                            XNew = labelPosition.LabelPoint.Lng - deltX;
                                            YNew = labelPosition.LabelPoint.Lat - deltY;
                                            labelPosition.LabelSouthWest = new Coord() { Lat = YNew, Lng = XNew - LabelWidth, Ordinal = 0 };
                                            labelPosition.LabelNorthEast = new Coord() { Lat = YNew - LabelHeight, Lng = XNew, Ordinal = 0 };
                                        }
                                        break;
                                    case PositionEnum.LeftTop:
                                        {
                                            XNew = labelPosition.LabelPoint.Lng + deltX;
                                            YNew = labelPosition.LabelPoint.Lat + deltY;
                                            labelPosition.LabelSouthWest = new Coord() { Lat = YNew + LabelHeight, Lng = XNew, Ordinal = 0 };
                                            labelPosition.LabelNorthEast = new Coord() { Lat = YNew, Lng = XNew + LabelWidth, Ordinal = 0 };
                                        }
                                        break;
                                    case PositionEnum.RightTop:
                                        {
                                            XNew = labelPosition.LabelPoint.Lng - deltX;
                                            YNew = labelPosition.LabelPoint.Lat + deltY;
                                            labelPosition.LabelSouthWest = new Coord() { Lat = YNew + LabelHeight, Lng = XNew - LabelWidth, Ordinal = 0 };
                                            labelPosition.LabelNorthEast = new Coord() { Lat = YNew, Lng = XNew, Ordinal = 0 };
                                        }
                                        break;
                                    default:
                                        break;
                                }
                                labelPosition.LabelPoint = new Coord() { Lat = YNew, Lng = XNew, Ordinal = 0 };
                                PleaseRedo = true;
                                break;
                            }
                        }
                        if (!PleaseRedo)
                        {
                            HidingPoint = false;
                        }
                    }

                    HidingPoint = true;
                    while (HidingPoint)
                    {
                        List<Coord> coordList = new List<Coord>()
                        {
                            new Coord() { Lat = labelPosition.LabelSouthWest.Lat, Lng = labelPosition.LabelSouthWest.Lng, Ordinal = 0 },
                            new Coord() { Lat = labelPosition.LabelSouthWest.Lat, Lng = labelPosition.LabelNorthEast.Lng, Ordinal = 0 },
                            new Coord() { Lat = labelPosition.LabelNorthEast.Lat, Lng = labelPosition.LabelNorthEast.Lng, Ordinal = 0 },
                            new Coord() { Lat = labelPosition.LabelNorthEast.Lat, Lng = labelPosition.LabelSouthWest.Lng, Ordinal = 0 },
                        };

                        bool PleaseRedo = false;
                        foreach (LabelPosition labelPosition2 in LabelPositionList.Where(c => c.Ordinal != labelPosition.Ordinal && c.Distance <= labelPosition.Distance))
                        {
                            List<Coord> coordToCompare = new List<Coord>()
                            {
                                new Coord() { Lat = labelPosition2.LabelSouthWest.Lat, Lng = labelPosition2.LabelSouthWest.Lng, Ordinal = 0 },
                                new Coord() { Lat = labelPosition2.LabelSouthWest.Lat, Lng = labelPosition2.LabelNorthEast.Lng, Ordinal = 0 },
                                new Coord() { Lat = labelPosition2.LabelNorthEast.Lat, Lng = labelPosition2.LabelNorthEast.Lng, Ordinal = 0 },
                                new Coord() { Lat = labelPosition2.LabelNorthEast.Lat, Lng = labelPosition2.LabelSouthWest.Lng, Ordinal = 0 },
                            };
                            for (int i = 0; i < 4; i++)
                            {
                                if (_MapInfoService.CoordInPolygon(coordList, coordToCompare[i]))
                                {
                                    float XNew = StepSize;
                                    float YNew = StepSize;
                                    float dist = (float)Math.Sqrt((AveragePoint.Y - labelPosition.SitePoint.Lat) * (AveragePoint.Y - labelPosition.SitePoint.Lat) + (AveragePoint.X - labelPosition.SitePoint.Lng) * (AveragePoint.X - labelPosition.SitePoint.Lng));
                                    float factor = dist / StepSize;
                                    float deltX = Math.Abs((AveragePoint.X - labelPosition.LabelPoint.Lng) / factor);
                                    float deltY = Math.Abs((AveragePoint.Y - labelPosition.LabelPoint.Lat) / factor);
                                    switch (labelPosition.Position)
                                    {
                                        case PositionEnum.Error:
                                            break;
                                        case PositionEnum.LeftBottom:
                                            {
                                                XNew = labelPosition.LabelPoint.Lng + deltX;
                                                YNew = labelPosition.LabelPoint.Lat - deltY;
                                                labelPosition.LabelSouthWest = new Coord() { Lat = YNew, Lng = XNew, Ordinal = 0 };
                                                labelPosition.LabelNorthEast = new Coord() { Lat = YNew - LabelHeight, Lng = XNew + LabelWidth, Ordinal = 0 };
                                            }
                                            break;
                                        case PositionEnum.RightBottom:
                                            {
                                                XNew = labelPosition.LabelPoint.Lng - deltX;
                                                YNew = labelPosition.LabelPoint.Lat - deltY;
                                                labelPosition.LabelSouthWest = new Coord() { Lat = YNew, Lng = XNew - LabelWidth, Ordinal = 0 };
                                                labelPosition.LabelNorthEast = new Coord() { Lat = YNew - LabelHeight, Lng = XNew, Ordinal = 0 };
                                            }
                                            break;
                                        case PositionEnum.LeftTop:
                                            {
                                                XNew = labelPosition.LabelPoint.Lng + deltX;
                                                YNew = labelPosition.LabelPoint.Lat + deltY;
                                                labelPosition.LabelSouthWest = new Coord() { Lat = YNew + LabelHeight, Lng = XNew, Ordinal = 0 };
                                                labelPosition.LabelNorthEast = new Coord() { Lat = YNew, Lng = XNew + LabelWidth, Ordinal = 0 };
                                            }
                                            break;
                                        case PositionEnum.RightTop:
                                            {
                                                XNew = labelPosition.LabelPoint.Lng - deltX;
                                                YNew = labelPosition.LabelPoint.Lat + deltY;
                                                labelPosition.LabelSouthWest = new Coord() { Lat = YNew + LabelHeight, Lng = XNew - LabelWidth, Ordinal = 0 };
                                                labelPosition.LabelNorthEast = new Coord() { Lat = YNew, Lng = XNew, Ordinal = 0 };
                                            }
                                            break;
                                        default:
                                            break;
                                    }
                                    labelPosition.LabelPoint = new Coord() { Lat = YNew, Lng = XNew, Ordinal = 0 };
                                    PleaseRedo = true;
                                    break;
                                }
                                if (PleaseRedo)
                                {
                                    break;
                                }
                            }
                        }
                        if (!PleaseRedo)
                        {
                            HidingPoint = false;
                        }
                    }
                }
                //using (Graphics g = pictureBoxLabelPosition.CreateGraphics())
                //{
                //    g.Clear(Color.White);

                //    foreach (LabelPosition labelPosition in LabelPositionList)
                //    {
                //        g.DrawLine(new Pen(Color.Blue, 1.0f), new PointF(labelPosition.SitePoint.Lng, labelPosition.SitePoint.Lat), AveragePoint);
                //        PointF p = new PointF();
                //        switch (labelPosition.Position)
                //        {
                //            case PositionEnum.Error:
                //                break;
                //            case PositionEnum.LeftBottom:
                //                {
                //                    p = new PointF(labelPosition.LabelSouthWest.Lng, labelPosition.LabelSouthWest.Lat);
                //                }
                //                break;
                //            case PositionEnum.RightBottom:
                //                {
                //                    p = new PointF(labelPosition.LabelNorthEast.Lng, labelPosition.LabelSouthWest.Lat);
                //                }
                //                break;
                //            case PositionEnum.LeftTop:
                //                {
                //                    p = new PointF(labelPosition.LabelSouthWest.Lng, labelPosition.LabelNorthEast.Lat);
                //                }
                //                break;
                //            case PositionEnum.RightTop:
                //                {
                //                    p = new PointF(labelPosition.LabelNorthEast.Lng, labelPosition.LabelNorthEast.Lat);
                //                }
                //                break;
                //            default:
                //                break;
                //        }
                //        g.DrawLine(new Pen(Color.Red, 1.0f), new PointF(labelPosition.SitePoint.Lng, labelPosition.SitePoint.Lat), p);
                //        g.DrawRectangle(new Pen(Color.Red, 1.0f), labelPosition.LabelSouthWest.Lng, labelPosition.LabelNorthEast.Lat, labelPosition.LabelNorthEast.Lng - labelPosition.LabelSouthWest.Lng, labelPosition.LabelSouthWest.Lat - labelPosition.LabelNorthEast.Lat);
                //    }
                //}
            }
        }
        private bool GetInset(Coord CoordNE, Coord CoordSW)
        {
            string NotUsed = "";
            double Lat = (CoordNE.Lat + CoordSW.Lat) / 2;
            double Lng = (CoordNE.Lng + CoordSW.Lng) / 2;

            GoogleLogoHeight = 24;
            int InsetZoom = 6;
            int InsetWidth = 200;
            int InsetHeight = 200;

            using (WebClient client = new WebClient())
            {
                string url = $@"https://maps.googleapis.com/maps/api/staticmap?maptype={ MapType }&center={ Lat.ToString("F6") },{ Lng.ToString("F6") }&zoom={ InsetZoom.ToString() }&size={ InsetWidth.ToString() }x{ InsetHeight.ToString() }&language={ LanguageRequest.ToString() }&key=AIzaSyAwPGpdSM6z0A7DFdWPbS3vIDTk2mxINaA";
                try
                {
                    client.DownloadFile(url, DirName + FileNameInset);
                }
                catch (Exception ex)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotDownloadWebAddress_Error_, url, ex.Message + ex.InnerException != null ? " Inner: " + ex.InnerException.Message : "");
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotDownloadWebAddress_Error_", url, ex.Message + ex.InnerException != null ? " Inner: " + ex.InnerException.Message : "");
                    return false;
                }
            }

            try
            {
                Coord coord = new Coord() { Lat = (float)Lat, Lng = (float)Lng, Ordinal = 0 };
                CoordMap coordMapTemp = _MapInfoService.GetBounds(coord, InsetZoom, InsetWidth, InsetHeight);

                Rectangle cropRect = new Rectangle(0, 0, InsetWidth, InsetHeight);
                using (Bitmap srcNewInset = new Bitmap(InsetWidth, InsetHeight))
                {
                    using (Graphics g = Graphics.FromImage(srcNewInset))
                    {
                        using (Bitmap srcInset = new Bitmap(DirName + FileNameInset))
                        {
                            g.DrawImage(srcInset, new Rectangle(0, 0, InsetWidth, InsetHeight), cropRect, GraphicsUnit.Pixel);
                        }

                        double LngX = 0.0D;
                        double LatY = 0.0D;
                        double TotalWidthLng = coordMapTemp.NorthEast.Lng - coordMapTemp.SouthWest.Lng;
                        double TotalHeightLat = coordMapTemp.NorthEast.Lat - coordMapTemp.SouthWest.Lat;

                        List<Point> polygonPointList = new List<Point>();
                        // point #1
                        LngX = ((CoordSW.Lng - coordMapTemp.SouthWest.Lng) / TotalWidthLng) * InsetWidth;
                        LatY = InsetHeight - ((TotalHeightLat - (coordMapTemp.NorthEast.Lat - CoordSW.Lat)) / TotalHeightLat) * InsetHeight;
                        polygonPointList.Add(new Point() { X = (int)LngX, Y = (int)LatY });
                        // point #2
                        LngX = ((CoordNE.Lng - coordMapTemp.SouthWest.Lng) / TotalWidthLng) * InsetWidth;
                        LatY = InsetHeight - ((TotalHeightLat - (coordMapTemp.NorthEast.Lat - CoordSW.Lat)) / TotalHeightLat) * InsetHeight;
                        polygonPointList.Add(new Point() { X = (int)LngX, Y = (int)LatY });
                        // point #3
                        LngX = ((CoordNE.Lng - coordMapTemp.SouthWest.Lng) / TotalWidthLng) * InsetWidth;
                        LatY = InsetHeight - ((TotalHeightLat - (coordMapTemp.NorthEast.Lat - CoordNE.Lat)) / TotalHeightLat) * InsetHeight;
                        polygonPointList.Add(new Point() { X = (int)LngX, Y = (int)LatY });
                        // point #4
                        LngX = ((CoordSW.Lng - coordMapTemp.SouthWest.Lng) / TotalWidthLng) * InsetWidth;
                        LatY = InsetHeight - ((TotalHeightLat - (coordMapTemp.NorthEast.Lat - CoordNE.Lat)) / TotalHeightLat) * InsetHeight;
                        polygonPointList.Add(new Point() { X = (int)LngX, Y = (int)LatY });
                        // point #5
                        LngX = ((CoordSW.Lng - coordMapTemp.SouthWest.Lng) / TotalWidthLng) * InsetWidth;
                        LatY = InsetHeight - ((TotalHeightLat - (coordMapTemp.NorthEast.Lat - CoordSW.Lat)) / TotalHeightLat) * InsetHeight;
                        polygonPointList.Add(new Point() { X = (int)LngX, Y = (int)LatY });

                        g.DrawPolygon(new Pen(Color.LightGreen, 1.0f), polygonPointList.ToArray());
                    }

                    srcNewInset.Save(DirName + FileNameInsetFinal, ImageFormat.Png);
                }


                cropRect = new Rectangle(0, 0, InsetWidth, InsetHeight - GoogleLogoHeight);

                using (Bitmap targetInsetFinal = new Bitmap(cropRect.Width, cropRect.Height))
                {
                    using (Graphics g = Graphics.FromImage(targetInsetFinal))
                    {
                        using (Bitmap srcInset = new Bitmap(DirName + FileNameInsetFinal))
                        {
                            g.DrawImage(srcInset, new Rectangle(0, 0, targetInsetFinal.Width, targetInsetFinal.Height), cropRect, GraphicsUnit.Pixel);

                            g.DrawRectangle(new Pen(Color.Black, 6.0f), cropRect);
                        }
                    }
                    targetInsetFinal.Save(DirName + FileNameInsetFinal, ImageFormat.Png);
                }
            }
            catch (Exception ex)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotCreate_ImageError_, TaskRunnerServiceRes.Inset, ex.Message + ex.InnerException != null ? " Inner: " + ex.InnerException.Message : "");
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotCreate_ImageError_", TaskRunnerServiceRes.Inset, ex.Message + ex.InnerException != null ? " Inner: " + ex.InnerException.Message : "");
                return false;
            }

            return true;
        }
        private CoordMap GetMapCoordinateWhileGettingGooglePNG(double MinLat, double MaxLat, double MinLng, double MaxLng)
        {
            CoordMap coordMap = new CoordMap();

            double ExtraLat = (MaxLat - MinLat) * 0.05D;
            double ExtraLng = (MaxLng - MinLng) * 0.05D;

            Coord coordNE = new Coord() { Lat = (float)(MaxLat + ExtraLat), Lng = (float)(MaxLng + ExtraLng), Ordinal = 0 };
            Coord coordSW = new Coord() { Lat = (float)(MinLat - ExtraLat), Lng = (float)(MinLng - ExtraLng), Ordinal = 0 };

            double CenterLat = coordNE.Lat - (coordNE.Lat - coordSW.Lat) / 4;
            double CenterLng = coordSW.Lng + (coordNE.Lng - coordSW.Lng) / 4;

            double deltaLng = 0.0D;
            double deltaLat = 0.0D;

            bool Found = false;
            bool ZoomWasReduced = false;
            while (!Found)
            {
                // calculate the center of the uper left block
                Coord coord = new Coord() { Lat = (float)CenterLat, Lng = (float)CenterLng, Ordinal = 0 };
                CoordMap coordMapTemp = _MapInfoService.GetBounds(coord, Zoom, GoogleImageWidth, GoogleImageHeight);
                deltaLng = Math.Abs((CenterLng - coordMapTemp.SouthWest.Lng) * 2);
                deltaLat = Math.Abs((CenterLat - coordMapTemp.SouthWest.Lat) * 2) - (Math.Abs((CenterLat - coordMapTemp.SouthWest.Lat) * 2) * GoogleLogoHeight / GoogleImageHeight);

                coordMap = _MapInfoService.GetBounds(new Coord() { Lat = (float)CenterLat, Lng = (float)CenterLng, Ordinal = 0 }, Zoom, GoogleImageWidth, GoogleImageHeight);
                CoordMap NewcoordMapTemp = _MapInfoService.GetBounds(new Coord() { Lat = (float)CenterLat, Lng = (float)(CenterLng + deltaLng), Ordinal = 0 }, Zoom, GoogleImageWidth, GoogleImageHeight);

                coordMap.NorthEast.Lng = NewcoordMapTemp.NorthEast.Lng;
                coordMap.NorthEast.Lat = NewcoordMapTemp.NorthEast.Lat;

                NewcoordMapTemp = _MapInfoService.GetBounds(new Coord() { Lat = (float)(CenterLat - deltaLat), Lng = (float)CenterLng, Ordinal = 0 }, Zoom, GoogleImageWidth, GoogleImageHeight);

                coordMap.SouthWest.Lng = NewcoordMapTemp.SouthWest.Lng;
                coordMap.SouthWest.Lat = NewcoordMapTemp.SouthWest.Lat;

                Found = true;
                if (coordMap.NorthEast.Lng < (double)coordNE.Lng)
                {
                    Zoom -= 1;
                    ZoomWasReduced = true;
                    Found = false;
                }
                if (coordMap.NorthEast.Lat < (double)coordNE.Lat)
                {
                    Zoom -= 1;
                    ZoomWasReduced = true;
                    Found = false;
                }
                if (coordMap.SouthWest.Lng > (double)coordSW.Lng)
                {
                    Zoom -= 1;
                    ZoomWasReduced = true;
                    Found = false;
                }
                if (coordMap.SouthWest.Lat > (double)coordSW.Lat)
                {
                    Zoom -= 1;
                    ZoomWasReduced = true;
                    Found = false;
                }
                if (Found && !ZoomWasReduced)
                {
                    if ((coordMap.NorthEast.Lng - coordMap.SouthWest.Lng) > (coordNE.Lng - coordSW.Lng) * 2)
                    {
                        Zoom += 1;
                        Found = false;
                    }
                    if ((coordMap.NorthEast.Lat - coordMap.SouthWest.Lat) > (coordNE.Lat - coordSW.Lat) * 2)
                    {
                        Zoom += 1;
                        Found = false;
                    }
                }
                if (Found)
                {
                    break;
                }
            }

            using (WebClient client = new WebClient())
            {
                try
                {
                    string url = $@"https://maps.googleapis.com/maps/api/staticmap?maptype={ MapType }&center={ CenterLat.ToString("F6") },{ CenterLng.ToString("F6") }&zoom={ Zoom.ToString() }&size={ GoogleImageWidth.ToString() }x{ GoogleImageHeight.ToString() }&language={ LanguageRequest.ToString() }&key=AIzaSyAwPGpdSM6z0A7DFdWPbS3vIDTk2mxINaA";
                    client.DownloadFile(url, DirName + FileNameNW);

                    url = $@"https://maps.googleapis.com/maps/api/staticmap?maptype={ MapType }&center={ CenterLat.ToString("F6") },{ (CenterLng + deltaLng).ToString("F6") }&zoom={ Zoom.ToString() }&size={ GoogleImageWidth.ToString() }x{ GoogleImageHeight.ToString() }&language={ LanguageRequest.ToString() }&key=AIzaSyAwPGpdSM6z0A7DFdWPbS3vIDTk2mxINaA";
                    client.DownloadFile(url, DirName + FileNameNE);

                    url = $@"https://maps.googleapis.com/maps/api/staticmap?maptype={ MapType }&center={ (CenterLat - deltaLat).ToString("F6") },{ CenterLng.ToString("F6") }&zoom={ Zoom.ToString() }&size={ GoogleImageWidth.ToString() }x{ GoogleImageHeight.ToString() }&language={ LanguageRequest.ToString() }&key=AIzaSyAwPGpdSM6z0A7DFdWPbS3vIDTk2mxINaA";
                    client.DownloadFile(url, DirName + FileNameSW);

                    url = $@"https://maps.googleapis.com/maps/api/staticmap?maptype={ MapType }&center={ (CenterLat - deltaLat).ToString("F6") },{ (CenterLng + deltaLng).ToString("F6") }&zoom={ Zoom.ToString() }&size={ GoogleImageWidth.ToString() }x{ GoogleImageHeight.ToString() }&language={ LanguageRequest.ToString() }&key=AIzaSyAwPGpdSM6z0A7DFdWPbS3vIDTk2mxINaA";
                    client.DownloadFile(url, DirName + FileNameSE);
                }
                catch (Exception ex)
                {
                    return new CoordMap();
                }
            }

            if (!CombineAllImageIntoOne())
            {
                return new CoordMap();
            }

            if (!DeleteTempGoogleImageFiles())
            {
                return new CoordMap();
            }

            return coordMap;
        }
        private void LoadDefaults()
        {
            Random random = new Random();
            FileNameExtra = "";
            for (int i = 0; i < 10; i++)
            {
                FileNameExtra = FileNameExtra + (char)random.Next(97, 122);
            }

            string ServerDir = _TVFileService.GetServerFilePath(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID);
            // defaults
            DirName = ServerDir;
            FileNameNW = $@"NW{ FileNameExtra }.png";
            FileNameNE = $@"NE{ FileNameExtra }.png";
            FileNameSW = $@"SW{ FileNameExtra }.png";
            FileNameSE = $@"SE{ FileNameExtra }.png";
            FileNameFull = $@"Full{ FileNameExtra }.png";
            FileNameInset = $@"Inset{ FileNameExtra }.png";
            FileNameInsetFinal = $@"InsetFinal{ FileNameExtra }.png";
            FileNameFullAnnotated = $@"Full{ FileNameExtra }Annotated.png";
            GoogleImageWidth = 640;
            GoogleImageHeight = 640;
            GoogleLogoHeight = 24;
            LanguageRequest = LanguageEnum.en;
            MapType = "hybrid"; // Can be roadmap (default), satellite, hybrid, terrain
            Zoom = 16;
        }
        #endregion Functions private
    }

}
