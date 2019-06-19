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
using System.Threading;

namespace CSSPWebToolsTaskRunner.Services
{
    public partial class ParametersService
    {
        private bool GenerateHTMLMikeScenario()
        {
            switch (reportTypeModel.UniqueCode)
            {
                case "MikeScenarioParametersAtOnePointXLSX":
                    {
                        if (!GenerateHTMLMikeScenario_MikeScenarioParametersAtOnePointXLSX())
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
        private bool GenerateHTMLMikeScenario_MikeScenarioParametersAtOnePointXLSX()
        {

            string NotUsed = "";
            bool ErrorInDoc = false;

            string Lat = "";
            string Lng = "";

            List<string> ParamValueList = Parameters.Split("|||".ToCharArray(), StringSplitOptions.RemoveEmptyEntries).ToList();

            Lat = GetParameters("Lat", ParamValueList);
            Lng = GetParameters("Lng", ParamValueList);

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
                    NotUsed = string.Format(TaskRunnerServiceRes.Error_WhileCreating_Document_, "GetHydrodynamicDfsuFile", "XLSX", fi.FullName);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("Error_WhileCreating_Document_", "GetHydrodynamicDfsuFile", "XLSX", fi.FullName);
                }
                return false;
            }

            dfsuFileTransport = GetTransportDfsuFile();
            if (dfsuFileTransport == null)
            {
                if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.Error_WhileCreating_Document_, "GetTransportDfsuFile", "XLSX", fi.FullName);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("Error_WhileCreating_Document_", "GetTransportDfsuFile", "XLSX", fi.FullName);
                }
                return false;
            }

            if (!FillRequiredList(dfsuFileTransport, elementList, elementLayerList, nodeList, topNodeLayerList, bottomNodeLayerList))
            {
                if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.Error_WhileCreating_Document_, "FillRequiredList", "XLSX", fi.FullName);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("Error_WhileCreating_Document_", "FillRequiredList", "XLSX", fi.FullName);
                }
                return false;
            }

            Node node = new Node();
            if (Thread.CurrentThread.CurrentCulture.Name == "fr-CA")
            {
                node.X = float.Parse(Lng.Replace(".", ","));
                node.Y = float.Parse(Lat.Replace(".", ","));
            }
            else
            {
                node.X = float.Parse(Lng);
                node.Y = float.Parse(Lat);
            }

            List<Node> NodeList = new List<Node>() { node };
            List<ElementLayer> SelectedElementLayerList = GetElementSurrondingEachPoint(elementLayerList, NodeList);
            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
                return false;

            if (SelectedElementLayerList.Count == 0)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.Error_WhileCreating_Document_, "FillRequiredList", "XLSX", fi.FullName);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("Error_WhileCreating_Document_", "FillRequiredList", "XLSX", fi.FullName);
                return false;
            }

            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
                return false;

            if (!GetTopHTML())
            {
                return false;
            }

            if (!WriteHTMLParameterContent(sb, dfsuFileHydrodynamic, dfsuFileTransport, elementLayerList, topNodeLayerList, bottomNodeLayerList, SelectedElementLayerList, node))
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
                NotUsed = string.Format(TaskRunnerServiceRes.ErrorOccuredWhileCreating_Document_, "XLSX", fi.FullName);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("ErrorOccuredWhileCreating_Document_", "XLSX", fi.FullName);
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
        private bool WriteHTMLParameterContent(StringBuilder sbHTML, DfsuFile dfsuFileHydrodynamic, DfsuFile dfsuFileTransport, List<ElementLayer> elementLayerList, List<NodeLayer> topNodeLayerList, List<NodeLayer> bottomNodeLayerList, List<ElementLayer> SelectedElementLayerList, Node node)
        {
            MapInfoService mapInfoService = new MapInfoService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);

            string NotUsed = "";

            int ItemUVelocity = 0;
            int ItemVVelocity = 0;
            int ItemSalinity = 0;
            int ItemTemperature = 0;
            int ItemWaterDepth = 0;

            if (SelectedElementLayerList.Count == 0)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.Error_WhileCreating_Document_, "FillRequiredList", "XLSX", fi.FullName);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("Error_WhileCreating_Document_", "FillRequiredList", "XLSX", fi.FullName);
                return false;
            }

            sb.AppendLine($@"<h2>Parameters output at Latitude: {node.Y} Longitude: {node.X}</h2>");

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

            foreach (IDfsSimpleDynamicItemInfo dfsDyInfo in dfsuFileHydrodynamic.ItemInfo)
            {
                if (dfsDyInfo.Quantity.Item == eumItem.eumISalinity)
                {
                    ItemSalinity = dfsDyInfo.ItemNumber;
                }
                if (dfsDyInfo.Quantity.Item == eumItem.eumITemperature)
                {
                    ItemTemperature = dfsDyInfo.ItemNumber;
                }
                if (dfsDyInfo.Quantity.Item == eumItem.eumIWaterDepth)
                {
                    ItemWaterDepth = dfsDyInfo.ItemNumber;
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
                //sbHTML.AppendLine(@"<th>Salinity</th>");
                //sbHTML.AppendLine(@"<th>Temperature</th>");
                sbHTML.AppendLine(@"<th>WaterDepth</th>");
                sbHTML.AppendLine(@"</tr>");
                sbHTML.AppendLine(@"</thead>");
                sbHTML.AppendLine(@"<tbody>");

                int vCount = 0;
                for (int timeStep = 0; timeStep < dfsuFileTransport.NumberOfTimeSteps; timeStep++)
                {
                    if (vCount % 10 == 0)
                    {
                        _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, (int)((float)(vCount / dfsuFileTransport.NumberOfTimeSteps) * 100.0f));
                    }

                    sbHTML.AppendLine(@"<tr>");
                    sbHTML.AppendLine($@"<td>{dfsuFileTransport.StartDateTime.AddSeconds(vCount * dfsuFileTransport.TimeStepInSeconds).ToString("yyyy-MM-dd HH:mm:ss")}</td>");
                    sbHTML.AppendLine($@"<td>{dfsuFileTransport.StartDateTime.AddSeconds((vCount + 1) * dfsuFileTransport.TimeStepInSeconds).ToString("yyyy-MM-dd HH:mm:ss")}</td>");

                    float[] UvelocityList = (float[])dfsuFileTransport.ReadItemTimeStep(ItemUVelocity, timeStep).Data;
                    float[] VvelocityList = (float[])dfsuFileTransport.ReadItemTimeStep(ItemVVelocity, timeStep).Data;

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

                    //string SalinityText = "Salinity";
                    //if (ItemSalinity != 0)
                    //{
                    //    float[] SalinityList = (float[])dfsuFileHydrodynamic.ReadItemTimeStep(ItemSalinity, timeStep).Data;

                    //    SalinityText = SalinityList[el.Element.ID - 1].ToString();
                    //}

                    //string TemperatureText = "Temperature";
                    //if (ItemTemperature != 0)
                    //{
                    //    float[] TemperatureList = (float[])dfsuFileHydrodynamic.ReadItemTimeStep(ItemTemperature, timeStep).Data;

                    //    TemperatureText = TemperatureList[el.Element.ID - 1].ToString();
                    //}

                    string WaterDepthText = "WaterDepth";
                    if (ItemWaterDepth != 0)
                    {
                        float[] TotalWaterDepthList = (float[])dfsuFileHydrodynamic.ReadItemTimeStep(ItemWaterDepth, timeStep).Data;

                        WaterDepthText = TotalWaterDepthList[el.Element.ID - 1].ToString();
                    }


                    sbHTML.AppendLine($@"<td>{UV}</td>");
                    sbHTML.AppendLine($@"<td>{VV}</td>");
                    sbHTML.AppendLine($@"<td>{VectorVal}</td>");
                    sbHTML.AppendLine($@"<td>{VectorDirCartesian}</td>");
                    //sbHTML.AppendLine($@"<td>{SalinityText}</td>");
                    //sbHTML.AppendLine($@"<td>{TemperatureText}</td>");
                    sbHTML.AppendLine($@"<td>{WaterDepthText}</td>");
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
