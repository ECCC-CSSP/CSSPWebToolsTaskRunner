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
using DHI.Generic.MikeZero.DFS.dfsu;
using DHI.Generic.MikeZero.DFS;
using DHI.Generic.MikeZero;

namespace CSSPWebToolsTaskRunner.Services
{
    public partial class ParametersService
    {
        private bool GenerateHTMLMikeScenario()
        {
            switch (reportTypeModel.UniqueCode)
            {
                case "MikeScenarioParameterAtOnePointXLSX":
                    {
                        if (!GenerateHTMLMikeScenario_MikeScenarioParameterAtOnePointXLSX())
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
                    if (!GenerateHTMLMikeScenario_NotImplemented())
                    {
                        return false;
                    }
                    break;
            }
            return true;
        }
        private bool GenerateHTMLMikeScenario_MikeScenarioParameterAtOnePointXLSX()
        {

            string NotUsed = "";
            bool ErrorInDoc = false;

            string GoogleEarthMarker = "";

            List<string> ParamValueList = Parameters.Split("|||".ToCharArray(), StringSplitOptions.RemoveEmptyEntries).ToList();

            GoogleEarthMarker = GetParameters("GoogleEarthMarker", ParamValueList);
            if (string.IsNullOrWhiteSpace(GoogleEarthMarker))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind__, TaskRunnerServiceRes.Parameter, TaskRunnerServiceRes.GoogleEarthMarker);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotFind__", TaskRunnerServiceRes.Parameter, TaskRunnerServiceRes.GoogleEarthMarker);
                return false;
            }

            GoogleEarthMarker = GoogleEarthMarker.Replace("!!!!!", "<");
            GoogleEarthMarker = GoogleEarthMarker.Replace("@@@@@", ">");
            GoogleEarthMarker = GoogleEarthMarker.Replace("%%%%%", ",");

            DfsuFile dfsuFileHydrodynamic = null;
            DfsuFile dfsuFileTransport = null;
            List<ElementLayer> elementLayerList = new List<ElementLayer>();
            List<Element> elementList = new List<Element>();
            List<Node> nodeList = new List<Node>();
            List<NodeLayer> topNodeLayerList = new List<NodeLayer>();
            List<NodeLayer> bottomNodeLayerList = new List<NodeLayer>();

            dfsuFileHydrodynamic = GetHydrodynamicDfsuFile();
            if (dfsuFileHydrodynamic == null)
            {
                if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.Error_WhileCreatingXLSXDocument_, "GetHydrodynamicDfsuFile", fi.FullName);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("Error_WhileCreatingXLSXDocument_", "GetHydrodynamicDfsuFile", fi.FullName);
                }
                return false;
            }

            dfsuFileTransport = GetTransportDfsuFile();
            if (dfsuFileTransport == null)
            {
                if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.Error_WhileCreatingXLSXDocument_, "GetTransportDfsuFile", fi.FullName);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("Error_WhileCreatingXLSXDocument_", "GetTransportDfsuFile", fi.FullName);
                }
                return false;
            }

            if (!FillRequiredList(dfsuFileTransport, elementList, elementLayerList, nodeList, topNodeLayerList, bottomNodeLayerList))
            {
                if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.Error_WhileCreatingXLSXDocument_, "FillRequiredList", fi.FullName);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("Error_WhileCreatingXLSXDocument_", "FillRequiredList", fi.FullName);
                }
                return false;
            }

            List<ElementLayer> SelectedElementLayerList = ParseKMLPath(GoogleEarthMarker, elementLayerList);
            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
                return false;

            if (SelectedElementLayerList.Count == 0)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.Error_WhileCreatingXLSXDocument_, "FillRequiredList", fi.FullName);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("Error_WhileCreatingXLSXDocument_", "FillRequiredList", fi.FullName);
                return false;
            }

            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
                return false;

            if (!GetTopHTML())
            {
                return false;
            }

            if (!WriteHTMLParameterContent(sb, dfsuFileHydrodynamic, dfsuFileTransport, elementLayerList, topNodeLayerList, bottomNodeLayerList, SelectedElementLayerList))
            {
                ErrorInDoc = true;
            }

            if (!GetBottomHTML())
            {
                return false;
            }

            dfsuFileHydrodynamic.Close();
            dfsuFileTransport.Close();

            if (ErrorInDoc)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.ErrorOccuredWhileCreatingXLSXDocument_, fi.FullName);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("ErrorOccuredWhileCreatingXLSXDocument_", fi.FullName);
                return false;
            }
            return true;
        }
        private bool GenerateHTMLMikeScenario_NotImplemented()
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
        private bool WriteHTMLParameterContent(StringBuilder sbHTML, DfsuFile dfsuFileHydrodynamic, DfsuFile dfsuFileTransport, List<ElementLayer> elementLayerList, List<NodeLayer> topNodeLayerList, List<NodeLayer> bottomNodeLayerList, List<ElementLayer> SelectedElementLayerList)
        {
            string NotUsed = "";

            int ItemUVelocity = 0;
            int ItemVVelocity = 0;
            int ItemSalinity = 0;
            int ItemTemperature = 0;
            int ItemSurfaceElevation = 0;
            int ItemTotalWaterDepth = 0;

            if (SelectedElementLayerList.Count == 0)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.Error_WhileCreatingXLSXDocument_, "FillRequiredList", fi.FullName);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("Error_WhileCreatingXLSXDocument_", "FillRequiredList", fi.FullName);
                return false;
            }

            Dictionary<int, Node> ElemCenter = new Dictionary<int, Node>();

            foreach (ElementLayer el in SelectedElementLayerList.Where(c => c.Layer == 1))
            {
                float XCenter = 0.0f;
                float YCenter = 0.0f;

                foreach (Node n in el.Element.NodeList)
                {
                    XCenter += n.X;
                    YCenter += n.Y;
                }
                XCenter = XCenter / el.Element.NodeList.Count();
                YCenter = YCenter / el.Element.NodeList.Count();

                ElemCenter.Add(el.Element.ID, new Node() { X = XCenter, Y = YCenter });
            }

            sb.AppendLine($@"<h2>Parameters output at Latitude: {ElemCenter[0].Y} Longitude: {ElemCenter[0].X}</h2>");

            // getting the ItemNumber
            foreach (IDfsSimpleDynamicItemInfo dfsDyInfo in dfsuFileTransport.ItemInfo)
            {
                if (dfsDyInfo.Quantity.Item == eumItem.eumIuVelocity)
                {
                    ItemUVelocity = dfsDyInfo.ItemNumber;
                }
                if (dfsDyInfo.Quantity.Item == eumItem.eumIvVelocity)
                {
                    ItemVVelocity = dfsDyInfo.ItemNumber;
                }
            }

            if (ItemUVelocity == 0 || ItemVVelocity == 0)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_, TaskRunnerServiceRes.Parameters);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotFind_", TaskRunnerServiceRes.Parameters);
                return false;
            }

            int CountLayer = (dfsuFileTransport.NumberOfSigmaLayers == 0 ? 1 : dfsuFileTransport.NumberOfSigmaLayers);

            for (int Layer = 1; Layer <= CountLayer; Layer++)
            {
                sb.AppendLine($@"<h2>Layer: {Layer}</h2>");

                sbHTML.AppendLine(@"<table>");
                sbHTML.AppendLine(@"<thead>");
                sbHTML.AppendLine(@"<tr>");
                sbHTML.AppendLine(@"<th>StartTime</th>");
                sbHTML.AppendLine(@"<th>EndTime</th>");
                sbHTML.AppendLine(@"<th>UVelocity</th>");
                sbHTML.AppendLine(@"<th>VVelocity</th>");
                sbHTML.AppendLine(@"<th>CurrentVelocity</th>");
                sbHTML.AppendLine(@"<th>CurrentDirection</th>");
                sbHTML.AppendLine(@"<th>Salinity</th>");
                sbHTML.AppendLine(@"<th>Temperature</th>");
                sbHTML.AppendLine(@"<th>SurfaceElevation</th>");
                sbHTML.AppendLine(@"<th>TotalWaterDepth</th>");
                sbHTML.AppendLine(@"</tr>");
                sbHTML.AppendLine(@"</thead>");
                sbHTML.AppendLine(@"<tbody>");

                ElemCenter = new Dictionary<int, Node>();

                foreach (ElementLayer el in SelectedElementLayerList.Where(c => c.Layer == Layer))
                {
                    float XCenter = 0.0f;
                    float YCenter = 0.0f;

                    foreach (Node n in el.Element.NodeList)
                    {
                        XCenter += n.X;
                        YCenter += n.Y;
                    }
                    XCenter = XCenter / el.Element.NodeList.Count();
                    YCenter = YCenter / el.Element.NodeList.Count();

                    ElemCenter.Add(el.Element.ID, new Node() { X = XCenter, Y = YCenter });
                }

                int vCount = 0;
                for (int timeStep = 0; timeStep < dfsuFileTransport.NumberOfTimeSteps; timeStep++)
                {
                    sbHTML.AppendLine(@"<tr>");
                    sbHTML.AppendLine($@"<td>{dfsuFileTransport.StartDateTime.AddSeconds(vCount * dfsuFileTransport.TimeStepInSeconds).ToString("yyyy-MM-dd HH:mm:ss")}</td>");
                    sbHTML.AppendLine($@"<td>{dfsuFileTransport.StartDateTime.AddSeconds((vCount + 1) * dfsuFileTransport.TimeStepInSeconds).ToString("yyyy-MM-dd HH:mm:ss")}</td>");



                    float[] UvelocityList = (float[])dfsuFileTransport.ReadItemTimeStep(ItemUVelocity, timeStep).Data;
                    float[] VvelocityList = (float[])dfsuFileTransport.ReadItemTimeStep(ItemVVelocity, timeStep).Data;

                    MapInfoService mapInfoService = new MapInfoService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
                    ElementLayer el = SelectedElementLayerList.Where(c => c.Layer == Layer).Take(1).FirstOrDefault();

                    float UV = UvelocityList[el.Element.ID - 1];
                    float VV = VvelocityList[el.Element.ID - 1];

                    double VectorVal = Math.Sqrt((UV * UV) + (VV * VV));
                    double VectorDir = 0.0D;
                    double VectorDirCartesian = Math.Acos(Math.Abs(UV / VectorVal)) * mapInfoService.r2d;

                    if (VectorDirCartesian <= 360 && VectorDirCartesian >= 0)
                    {
                        // everything is ok
                    }
                    else
                    {
                        VectorDirCartesian = 0.0D;
                    }

                    if (UV >= 0 && VV >= 0)
                    {
                        VectorDir = 90 - VectorDirCartesian;
                    }
                    else if (UV < 0 && VV >= 0)
                    {
                        VectorDir = 270 + VectorDirCartesian;
                        VectorDirCartesian = 180 - VectorDirCartesian;
                    }
                    else if (UV >= 0 && VV < 0)
                    {
                        VectorDir = 90 + VectorDirCartesian;
                        VectorDirCartesian = 360 - VectorDirCartesian;
                    }
                    else if (UV < 0 && VV < 0)
                    {
                        VectorDir = 270 - VectorDirCartesian;
                        VectorDirCartesian = 180 + VectorDirCartesian;
                    }

                    sbHTML.AppendLine($@"<td>{UV}</td>");
                    sbHTML.AppendLine($@"<td>{VV}</td>");
                    sbHTML.AppendLine($@"<td>{VectorVal}</td>");
                    sbHTML.AppendLine($@"<td>{VectorDirCartesian}</td>");
                    sbHTML.AppendLine($@"<td>Salinity</td>");
                    sbHTML.AppendLine($@"<td>Temperature</td>");
                    sbHTML.AppendLine($@"<td>SurfaceElevation</td>");
                    sbHTML.AppendLine($@"<td>TotalWaterDepth</td>");
                    sbHTML.AppendLine(@"</tr>");

                    vCount += 1;
                }

                sbHTML.AppendLine(@"</tbody>");
                sbHTML.AppendLine(@"</table>");

            }

            return true;
        }

    }
}
