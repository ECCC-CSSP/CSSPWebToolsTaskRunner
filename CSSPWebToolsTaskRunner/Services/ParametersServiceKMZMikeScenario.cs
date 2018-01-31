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
using System.Threading;
using System.Globalization;
using DHI.Generic.MikeZero.DFS.dfsu;
using DHI.PFS;
using DHI.Generic.MikeZero.DFS;
using DHI.Generic.MikeZero;

namespace CSSPWebToolsTaskRunner.Services
{
    public partial class ParametersService
    {
        // commands
        private bool GenerateKMZMikeScenario()
        {
            switch (reportTypeModel.UniqueCode)
            {
                case "MikeScenarioBoundaryConditionsKMZ":
                    {
                        if (!GenerateMikeScenarioBoundaryConditionsKMZ())
                        {
                            return false;
                        }
                    }
                    break;
                case "MikeScenarioMeshKMZ":
                    {
                        if (!GenerateMikeScenarioMeshKMZ())
                        {
                            return false;
                        }
                    }
                    break;
                case "MikeScenarioPollutionLimitKMZ":
                    {
                        if (!GenerateMikeScenarioPollutionLimitKMZ())
                        {
                            return false;
                        }
                    }
                    break;
                case "MikeScenarioPollutionAnimationKMZ":
                    {
                        if (!GenerateMikeScenarioPollutionAnimationKMZ())
                        {
                            return false;
                        }
                    }
                    break;
                case "MikeScenarioStudyAreaKMZ":
                    {
                        if (!GenerateMikeScenarioStudyAreaKMZ())
                        {
                            return false;
                        }
                    }
                    break;
                case "MikeScenarioTestKMZ":
                    {
                        if (!GenerateKMZMikeScenario_MikeScenarioTestKMZ())
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
                    break;
            }
            return true;
        }
        private bool GenerateMikeScenarioBoundaryConditionsKMZ()
        {
            string NotUsed = "";

            bool ErrorInDoc = false;
            if (!WriteKMLTop(sb))
            {
                ErrorInDoc = true;
            }
            sb.AppendLine(string.Format(@"<name>{0}</name>", TaskRunnerServiceRes.MIKEBoundaryConditions + "_" + GetScenarioName()));
            if (!WriteKMLBoundaryConditionNode(sb))
            {
                ErrorInDoc = true;
            }

            if (!WriteKMLBottom(sb))
            {
                ErrorInDoc = true;
            }

            if (ErrorInDoc)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.ErrorOccuredWhileCreatingKMZDocument_, fi.FullName);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("ErrorOccuredWhileCreatingKMZDocument_", fi.FullName);
                return false;
            }

            return true;
        }
        private bool GenerateMikeScenarioMeshKMZ()
        {
            string NotUsed = "";
            bool ErrorInDoc = false;

            DfsuFile dfsuFile = null;
            List<ElementLayer> elementLayerList = new List<ElementLayer>();
            List<Element> elementList = new List<Element>();
            List<Node> nodeList = new List<Node>();
            List<NodeLayer> topNodeLayerList = new List<NodeLayer>();
            List<NodeLayer> bottomNodeLayerList = new List<NodeLayer>();

            dfsuFile = GetHydrodynamicDfsuFile();
            if (dfsuFile == null)
            {
                if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.Error_WhileCreatingKMZDocument_, "GettingHydrodynamicDfsuFile", fi.FullName);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("Error_WhileCreatingKMZDocument_", "GettingHydrodynamicDfsuFile", fi.FullName);
                }
                return false;
            }
            if (!FillRequiredList(dfsuFile, elementLayerList, elementList, nodeList, topNodeLayerList, bottomNodeLayerList))
            {
                if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.Error_WhileCreatingKMZDocument_, "FillRequiredList", fi.FullName);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("Error_WhileCreatingKMZDocument_", "FillRequiredList", fi.FullName);
                }
                return false;
            }

            if (!WriteKMLTop(sb))
            {
                ErrorInDoc = true;
            }
            sb.AppendLine(string.Format(@"<name>{0}</name>", TaskRunnerServiceRes.MIKEMesh + "_" + GetScenarioName()));
            if (!WriteKMLMesh(sb, elementLayerList))
            {
                ErrorInDoc = true;
            }

            if (!WriteKMLBottom(sb))
            {
                ErrorInDoc = true;
            }

            dfsuFile.Close();

            if (ErrorInDoc)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.ErrorOccuredWhileCreatingKMZDocument_, fi.FullName);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("ErrorOccuredWhileCreatingKMZDocument_", fi.FullName);
                return false;
            }

            return true;
        }
        private bool GenerateMikeScenarioStudyAreaKMZ()
        {
            string NotUsed = "";
            bool ErrorInDoc = false;

            DfsuFile dfsuFile = null;
            List<ElementLayer> elementLayerList = new List<ElementLayer>();
            List<Element> elementList = new List<Element>();
            List<Node> nodeList = new List<Node>();
            List<NodeLayer> topNodeLayerList = new List<NodeLayer>();
            List<NodeLayer> bottomNodeLayerList = new List<NodeLayer>();

            dfsuFile = GetHydrodynamicDfsuFile();
            if (dfsuFile == null)
            {
                if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.Error_WhileCreatingKMZDocument_, "GettingHydrodynamicDfsuFile", fi.FullName);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("Error_WhileCreatingKMZDocument_", "GettingHydrodynamicDfsuFile", fi.FullName);
                }
                return false;
            }
            if (!FillRequiredList(dfsuFile, elementLayerList, elementList, nodeList, topNodeLayerList, bottomNodeLayerList))
            {
                if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.Error_WhileCreatingKMZDocument_, "FillRequiredList", fi.FullName);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("Error_WhileCreatingKMZDocument_", "FillRequiredList", fi.FullName);
                }
                return false;
            }

            if (!WriteKMLTop(sb))
            {
                ErrorInDoc = true;
            }
            sb.AppendLine(string.Format(@"<name>{0}</name>", TaskRunnerServiceRes.MIKEStudyArea + "_" + GetScenarioName()));
            if (!WriteKMLStudyAreaLine(sb, elementLayerList, nodeList))
            {
                ErrorInDoc = true;
            }

            if (!WriteKMLBottom(sb))
            {
                ErrorInDoc = true;
            }

            dfsuFile.Close();

            if (ErrorInDoc)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.ErrorOccuredWhileCreatingKMZDocument_, fi.FullName);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("ErrorOccuredWhileCreatingKMZDocument_", fi.FullName);
                return false;
            }

            return true;
        }
        private bool GenerateMikeScenarioPollutionAnimationKMZ()
        {
            string NotUsed = "";
            bool ErrorInDoc = false;

            string ContourValues = "";
            List<string> ParamValueList = Parameters.Split("|||".ToCharArray(), StringSplitOptions.RemoveEmptyEntries).ToList();

            ContourValues = GetParameters("ContourValues", ParamValueList);
            if (string.IsNullOrWhiteSpace(ContourValues))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind__, TaskRunnerServiceRes.Parameter, TaskRunnerServiceRes.ContourValues);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotFind__", TaskRunnerServiceRes.Parameter, TaskRunnerServiceRes.ContourValues);
                return false;
            }

            List<float> ContourValueList = ContourValues.Split(" ".ToCharArray(), StringSplitOptions.RemoveEmptyEntries).Select(c => float.Parse(c)).ToList();

            DfsuFile dfsuFile = null;
            List<ElementLayer> elementLayerList = new List<ElementLayer>();
            List<Element> elementList = new List<Element>();
            List<Node> nodeList = new List<Node>();
            List<NodeLayer> topNodeLayerList = new List<NodeLayer>();
            List<NodeLayer> bottomNodeLayerList = new List<NodeLayer>();

            dfsuFile = GetTransportDfsuFile();
            if (dfsuFile == null)
            {
                if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.Error_WhileCreatingKMZDocument_, "GetTransportDfsuFile", fi.FullName);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("Error_WhileCreatingKMZDocument_", "GetTransportDfsuFile", fi.FullName);
                }
                return false;
            }
            if (!FillRequiredList(dfsuFile, elementLayerList, elementList, nodeList, topNodeLayerList, bottomNodeLayerList))
            {
                if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.Error_WhileCreatingKMZDocument_, "FillRequiredList", fi.FullName);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("Error_WhileCreatingKMZDocument_", "FillRequiredList", fi.FullName);
                }
                return false;
            }

            if (!WriteKMLTop(sb))
            {
                ErrorInDoc = true;
            }
            sb.AppendLine(string.Format(@"<name>{0}</name>", TaskRunnerServiceRes.MIKEPollutionAnimation + "_" + GetScenarioName()));

            if (!WriteKMLStyleModelInput(sb))
            {
                ErrorInDoc = true;
            }

            if (!DrawKMLContourStyle(sb))
            {
                ErrorInDoc = true;
            }

            if (!WriteKMLModelInput(sb, ContourValueList))
            {
                ErrorInDoc = true;
            }
            if (!WriteKMLFecalColiformContourLine(dfsuFile, elementLayerList, topNodeLayerList, bottomNodeLayerList, ContourValueList))
            {
                ErrorInDoc = true;
            }

            if (!WriteKMLBottom(sb))
            {
                ErrorInDoc = true;
            }

            dfsuFile.Close();

            if (ErrorInDoc)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.ErrorOccuredWhileCreatingKMZDocument_, fi.FullName);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("ErrorOccuredWhileCreatingKMZDocument_", fi.FullName);
                return false;
            }
            return true;
        }
        private bool GenerateMikeScenarioPollutionLimitKMZ()
        {
            string NotUsed = "";
            bool ErrorInDoc = false;

            string ContourValues = "";
            List<string> ParamValueList = Parameters.Split("|||".ToCharArray(), StringSplitOptions.RemoveEmptyEntries).ToList();

            ContourValues = GetParameters("ContourValues", ParamValueList);
            if (string.IsNullOrWhiteSpace(ContourValues))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind__, TaskRunnerServiceRes.Parameter, TaskRunnerServiceRes.ContourValues);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotFind__", TaskRunnerServiceRes.Parameter, TaskRunnerServiceRes.ContourValues);
                return false;
            }

            List<float> ContourValueList = ContourValues.Split(" ".ToCharArray(), StringSplitOptions.RemoveEmptyEntries).Select(c => float.Parse(c)).ToList();

            DfsuFile dfsuFile = null;
            List<ElementLayer> elementLayerList = new List<ElementLayer>();
            List<Element> elementList = new List<Element>();
            List<Node> nodeList = new List<Node>();
            List<NodeLayer> topNodeLayerList = new List<NodeLayer>();
            List<NodeLayer> bottomNodeLayerList = new List<NodeLayer>();

            dfsuFile = GetTransportDfsuFile();
            if (dfsuFile == null)
            {
                if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.Error_WhileCreatingKMZDocument_, "GetTransportDfsuFile", fi.FullName);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("Error_WhileCreatingKMZDocument_", "GetTransportDfsuFile", fi.FullName);
                }
                return false;
            }
            if (!FillRequiredList(dfsuFile, elementLayerList, elementList, nodeList, topNodeLayerList, bottomNodeLayerList))
            {
                if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.Error_WhileCreatingKMZDocument_, "FillRequiredList", fi.FullName);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("Error_WhileCreatingKMZDocument_", "FillRequiredList", fi.FullName);
                }
                return false;
            }

            if (!WriteKMLTop(sb))
            {
                ErrorInDoc = true;
            }

            sb.AppendLine(string.Format(@"<name>{0}</name>", TaskRunnerServiceRes.MIKEPollutionLimit + "_" + GetScenarioName()));

            if (!WriteKMLStyleModelInput(sb))
            {
                ErrorInDoc = true;
            }

            if (!DrawKMLContourStyle(sb))
            {
                ErrorInDoc = true;
            }

            if (!WriteKMLModelInput(sb, ContourValueList))
            {
                ErrorInDoc = true;
            }
            if (!WriteKMLPollutionLimitsContourLine(dfsuFile, sb, elementLayerList, topNodeLayerList, bottomNodeLayerList, ContourValueList))
            {
                ErrorInDoc = true;
            }

            if (!WriteKMLBottom(sb))
            {
                ErrorInDoc = true;
            }

            dfsuFile.Close();

            if (ErrorInDoc)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.ErrorOccuredWhileCreatingKMZDocument_, fi.FullName);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("ErrorOccuredWhileCreatingKMZDocument_", fi.FullName);
                return false;
            }

            return true;
        }
        private bool GenerateKMZMikeScenario_MikeScenarioTestKMZ()
        {
            sb.AppendLine(@"<?xml version=""1.0"" encoding=""UTF-8""?>");
            sb.AppendLine(@"<kml xmlns=""http://www.opengis.net/kml/2.2"" xmlns:gx=""http://www.google.com/kml/ext/2.2"" xmlns:kml=""http://www.opengis.net/kml/2.2"" xmlns:atom=""http://www.w3.org/2005/Atom"">");
            sb.AppendLine(@"<Document>");
            sb.AppendLine(@"	<name>" + fi.FullName.Replace(".kml", ".kmz") + "</name>");
            sb.AppendLine(@"	<Placemark>");
            sb.AppendLine(@"		<name>My Parents Home</name> ");
            sb.AppendLine(@"		<Point>");
            sb.AppendLine(@"			<coordinates>-64.69002452357093,46.48465663502946,0</coordinates>");
            sb.AppendLine(@"		</Point> ");
            sb.AppendLine(@"	</Placemark>");
            sb.AppendLine(@"</Document> ");
            sb.AppendLine(@"</kml>");

            return true;
        }
        private bool GenerateKMZMikeScenario_NotImplementedKMZ(FileInfo fi, StringBuilder sbHTML, string parameters, ReportTypeModel reportTypeModel)
        {
            sbHTML.AppendLine(@"<?xml version=""1.0"" encoding=""UTF-8""?>");
            sbHTML.AppendLine(@"<kml xmlns=""http://www.opengis.net/kml/2.2"" xmlns:gx=""http://www.google.com/kml/ext/2.2"" xmlns:kml=""http://www.opengis.net/kml/2.2"" xmlns:atom=""http://www.w3.org/2005/Atom"">");
            sbHTML.AppendLine(@"<Document>");
            sbHTML.AppendLine(@"	<name>" + fi.FullName.Replace(".kml", ".kmz") + "</name>");
            sbHTML.AppendLine(@"	<Placemark>");
            sbHTML.AppendLine(@"		<name>Not Implemented</name> ");
            sbHTML.AppendLine(@"		<Point>");
            sbHTML.AppendLine(@"			<coordinates>-90,50,0</coordinates>");
            sbHTML.AppendLine(@"		</Point> ");
            sbHTML.AppendLine(@"	</Placemark>");
            sbHTML.AppendLine(@"</Document> ");
            sbHTML.AppendLine(@"</kml>");

            return true;
        }


        // helpers
        private void DrawKMLContourPolygon(List<ContourPolygon> ContourPolygonList, DfsuFile dfsuFile, int ParamCount)
        {
            int Count = 0;
            float MaxXCoord = -180;
            float MaxYCoord = -90;
            float MinXCoord = 180;
            float MinYCoord = 90;
            sb.AppendLine(@"  <Folder>");
            sb.AppendLine(@"    <visibility>0</visibility>");
            sb.AppendLine(string.Format(@"    <name>{0:yyyy-MM-dd} {0:HH:mm:ss tt}</name>", dfsuFile.StartDateTime.AddSeconds(ParamCount * dfsuFile.TimeStepInSeconds)));
            sb.AppendLine(@"    <TimeSpan>");
            sb.AppendLine(string.Format(@"    <begin>{0:yyyy-MM-dd}T{0:HH:mm:ss}</begin>", dfsuFile.StartDateTime.AddSeconds(ParamCount * dfsuFile.TimeStepInSeconds)));
            sb.AppendLine(string.Format(@"    <end>{0:yyyy-MM-dd}T{0:HH:mm:ss}</end>", dfsuFile.StartDateTime.AddSeconds((ParamCount + 1) * dfsuFile.TimeStepInSeconds)));
            sb.AppendLine(@"    </TimeSpan>");
            foreach (ContourPolygon contourPolygon in ContourPolygonList)
            {
                Count += 1;
                // draw the polygons
                sb.AppendLine(@"    <Placemark>");
                sb.AppendLine(@"      <visibility>0</visibility>");
                sb.AppendLine(string.Format(@"      <name>{0} {1}</name>", contourPolygon.ContourValue, TaskRunnerServiceRes.PollutionContour));
                if (contourPolygon.ContourValue >= 14 && contourPolygon.ContourValue < 88)
                {
                    sb.AppendLine(@"      <styleUrl>#fc_14</styleUrl>");
                }
                else if (contourPolygon.ContourValue >= 88)
                {
                    sb.AppendLine(@"      <styleUrl>#fc_88</styleUrl>");
                }
                else
                {
                    sb.AppendLine(@"      <styleUrl>#fc_LT14</styleUrl>");
                }

                sb.AppendLine(@"        <Polygon>");
                //sbPlacemarkFeacalColiformContour.AppendLine(@"<extrude>1</extrude>");
                //sbPlacemarkFeacalColiformContour.AppendLine(@"<tessellate>1</tessellate>");
                //sbPlacemarkFeacalColiformContour.AppendLine(@"<altitudeMode>absolute</altitudeMode>");
                //sbPlacemarkFeacalColiformContour.AppendLine(@"<gx:drawOrder>" + contourPolygon.Layer +"</gx:drawOrder>"); 
                sb.AppendLine(@"        <outerBoundaryIs>");
                sb.AppendLine(@"        <LinearRing>");
                sb.AppendLine(@"        <coordinates>");
                foreach (Node node in contourPolygon.ContourNodeList)
                {
                    if (MaxXCoord < node.X) MaxXCoord = node.X;
                    if (MaxYCoord < node.Y) MaxYCoord = node.Y;
                    if (MinXCoord > node.X) MinXCoord = node.X;
                    if (MinYCoord > node.Y) MinYCoord = node.Y;
                    sb.Append("        " + node.X.ToString().Replace(",", ".") + @"," + node.Y.ToString().Replace(",", ".") + "," + node.Z.ToString().Replace(",", ".") + " ");
                }
                sb.AppendLine(@"        </coordinates>");
                sb.AppendLine(@"        </LinearRing>");
                sb.AppendLine(@"        </outerBoundaryIs>");
                sb.AppendLine(@"        </Polygon>");
                sb.AppendLine(@"    </Placemark>");
            }
            sb.AppendLine(@"  </Folder>");
        }
        private bool DrawKMLContourStyle(StringBuilder sbHTML)
        {
            sbHTML.AppendLine(@"	<StyleMap id=""msn_ylw-pushpin"">");
            sbHTML.AppendLine(@"		<Pair>");
            sbHTML.AppendLine(@"			<key>normal</key>");
            sbHTML.AppendLine(@"			<styleUrl>#sn_ylw-pushpin</styleUrl>");
            sbHTML.AppendLine(@"		</Pair>");
            sbHTML.AppendLine(@"		<Pair>");
            sbHTML.AppendLine(@"			<key>highlight</key>");
            sbHTML.AppendLine(@"			<styleUrl>#sh_ylw-pushpin</styleUrl>");
            sbHTML.AppendLine(@"		</Pair>");
            sbHTML.AppendLine(@"	</StyleMap>");
            sbHTML.AppendLine(@"	<Style id=""sn_ylw-pushpin"">");
            sbHTML.AppendLine(@"		<IconStyle>");
            sbHTML.AppendLine(@"			<scale>1.1</scale>");
            sbHTML.AppendLine(@"			<Icon>");
            sbHTML.AppendLine(@"				<href>http://maps.google.com/mapfiles/kml/pushpin/ylw-pushpin.png</href>");
            sbHTML.AppendLine(@"			</Icon>");
            sbHTML.AppendLine(@"			<hotSpot x=""20"" y=""2"" xunits=""pixels"" yunits=""pixels""/>");
            sbHTML.AppendLine(@"		</IconStyle>");
            sbHTML.AppendLine(@"      <LineStyle>");
            sbHTML.AppendLine(@"         <color>ff000000</color>");
            sbHTML.AppendLine(@"       </LineStyle>");
            sbHTML.AppendLine(@"	</Style>");
            sbHTML.AppendLine(@"	<Style id=""sh_ylw-pushpin"">");
            sbHTML.AppendLine(@"		<IconStyle>");
            sbHTML.AppendLine(@"			<scale>1.3</scale>");
            sbHTML.AppendLine(@"			<Icon>");
            sbHTML.AppendLine(@"				<href>http://maps.google.com/mapfiles/kml/pushpin/ylw-pushpin.png</href>");
            sbHTML.AppendLine(@"			</Icon>");
            sbHTML.AppendLine(@"			<hotSpot x=""20"" y=""2"" xunits=""pixels"" yunits=""pixels""/>");
            sbHTML.AppendLine(@"		</IconStyle>");
            sbHTML.AppendLine(@"      <LineStyle>");
            sbHTML.AppendLine(@"         <color>ff000000</color>");
            sbHTML.AppendLine(@"       </LineStyle>");
            sbHTML.AppendLine(@"	</Style>");

            sbHTML.AppendLine(@"	<StyleMap id=""msn_grn-pushpin"">");
            sbHTML.AppendLine(@"		<Pair>");
            sbHTML.AppendLine(@"			<key>normal</key>");
            sbHTML.AppendLine(@"			<styleUrl>#sn_grn-pushpin</styleUrl>");
            sbHTML.AppendLine(@"		</Pair>");
            sbHTML.AppendLine(@"		<Pair>");
            sbHTML.AppendLine(@"			<key>highlight</key>");
            sbHTML.AppendLine(@"			<styleUrl>#sh_grn-pushpin</styleUrl>");
            sbHTML.AppendLine(@"		</Pair>");
            sbHTML.AppendLine(@"	</StyleMap>");
            sbHTML.AppendLine(@"	<Style id=""sn_grn-pushpin"">");
            sbHTML.AppendLine(@"		<IconStyle>");
            sbHTML.AppendLine(@"			<scale>1.1</scale>");
            sbHTML.AppendLine(@"			<Icon>");
            sbHTML.AppendLine(@"				<href>http://maps.google.com/mapfiles/kml/pushpin/grn-pushpin.png</href>");
            sbHTML.AppendLine(@"			</Icon>");
            sbHTML.AppendLine(@"			<hotSpot x=""20"" y=""2"" xunits=""pixels"" yunits=""pixels""/>");
            sbHTML.AppendLine(@"		</IconStyle>");
            sbHTML.AppendLine(@"      <LineStyle>");
            sbHTML.AppendLine(@"         <color>ff000000</color>");
            sbHTML.AppendLine(@"       </LineStyle>");
            sbHTML.AppendLine(@"	</Style>");
            sbHTML.AppendLine(@"	<Style id=""sh_grn-pushpin"">");
            sbHTML.AppendLine(@"		<IconStyle>");
            sbHTML.AppendLine(@"			<scale>1.3</scale>");
            sbHTML.AppendLine(@"			<Icon>");
            sbHTML.AppendLine(@"				<href>http://maps.google.com/mapfiles/kml/pushpin/grn-pushpin.png</href>");
            sbHTML.AppendLine(@"			</Icon>");
            sbHTML.AppendLine(@"			<hotSpot x=""20"" y=""2"" xunits=""pixels"" yunits=""pixels""/>");
            sbHTML.AppendLine(@"		</IconStyle>");
            sbHTML.AppendLine(@"      <LineStyle>");
            sbHTML.AppendLine(@"         <color>ff000000</color>");
            sbHTML.AppendLine(@"       </LineStyle>");
            sbHTML.AppendLine(@"	</Style>");

            sbHTML.AppendLine(@"	<StyleMap id=""msn_blue-pushpin"">");
            sbHTML.AppendLine(@"		<Pair>");
            sbHTML.AppendLine(@"			<key>normal</key>");
            sbHTML.AppendLine(@"			<styleUrl>#sn_blue-pushpin</styleUrl>");
            sbHTML.AppendLine(@"		</Pair>");
            sbHTML.AppendLine(@"		<Pair>");
            sbHTML.AppendLine(@"			<key>highlight</key>");
            sbHTML.AppendLine(@"			<styleUrl>#sh_blue-pushpin</styleUrl>");
            sbHTML.AppendLine(@"		</Pair>");
            sbHTML.AppendLine(@"	</StyleMap>");
            sbHTML.AppendLine(@"	<Style id=""sn_blue-pushpin"">");
            sbHTML.AppendLine(@"		<IconStyle>");
            sbHTML.AppendLine(@"			<scale>1.1</scale>");
            sbHTML.AppendLine(@"			<Icon>");
            sbHTML.AppendLine(@"				<href>http://maps.google.com/mapfiles/kml/pushpin/blue-pushpin.png</href>");
            sbHTML.AppendLine(@"			</Icon>");
            sbHTML.AppendLine(@"			<hotSpot x=""20"" y=""2"" xunits=""pixels"" yunits=""pixels""/>");
            sbHTML.AppendLine(@"		</IconStyle>");
            sbHTML.AppendLine(@"      <LineStyle>");
            sbHTML.AppendLine(@"         <color>ff000000</color>");
            sbHTML.AppendLine(@"       </LineStyle>");
            sbHTML.AppendLine(@"	</Style>");
            sbHTML.AppendLine(@"	<Style id=""sh_blue-pushpin"">");
            sbHTML.AppendLine(@"		<IconStyle>");
            sbHTML.AppendLine(@"			<scale>1.3</scale>");
            sbHTML.AppendLine(@"			<Icon>");
            sbHTML.AppendLine(@"				<href>http://maps.google.com/mapfiles/kml/pushpin/blue-pushpin.png</href>");
            sbHTML.AppendLine(@"			</Icon>");
            sbHTML.AppendLine(@"			<hotSpot x=""20"" y=""2"" xunits=""pixels"" yunits=""pixels""/>");
            sbHTML.AppendLine(@"		</IconStyle>");
            sbHTML.AppendLine(@"      <LineStyle>");
            sbHTML.AppendLine(@"         <color>ff000000</color>");
            sbHTML.AppendLine(@"       </LineStyle>");
            sbHTML.AppendLine(@"	</Style>");

            sbHTML.AppendLine(@"<Style id=""fc_LT14"">");
            sbHTML.AppendLine(@"<LineStyle>");
            sbHTML.AppendLine(@"<color>ff000000</color>");
            sbHTML.AppendLine(@"</LineStyle>");
            sbHTML.AppendLine(@"<PolyStyle>");
            sbHTML.AppendLine(@"<color>6600ff00</color>");
            sbHTML.AppendLine(@"<outline>1</outline>");
            sbHTML.AppendLine(@"</PolyStyle>");
            sbHTML.AppendLine(@"</Style>");

            sbHTML.AppendLine(@"<Style id=""fc_14"">");
            sbHTML.AppendLine(@"<LineStyle>");
            sbHTML.AppendLine(@"<color>ff000000</color>");
            sbHTML.AppendLine(@"</LineStyle>");
            sbHTML.AppendLine(@"<PolyStyle>");
            sbHTML.AppendLine(@"<color>66ff0000</color>");
            sbHTML.AppendLine(@"<outline>1</outline>");
            sbHTML.AppendLine(@"</PolyStyle>");
            sbHTML.AppendLine(@"</Style>");

            sbHTML.AppendLine(@"<Style id=""fc_88"">");
            sbHTML.AppendLine(@"<LineStyle>");
            sbHTML.AppendLine(@"<color>ff000000</color>");
            sbHTML.AppendLine(@"</LineStyle>");
            sbHTML.AppendLine(@"<PolyStyle>");
            sbHTML.AppendLine(@"<color>660000ff</color>");
            sbHTML.AppendLine(@"<outline>1</outline>");
            sbHTML.AppendLine(@"</PolyStyle>");
            sbHTML.AppendLine(@"</Style>");

            return true;
        }
        public bool FillElementLayerList(DfsuFile dfsuFile, List<Element> ElementList, List<ElementLayer> ElementLayerList, List<NodeLayer> topNodeLayerList, List<NodeLayer> bottomNodeLayerList)
        {
            string NotUsed = "";

            if (dfsuFile.DfsuFileType == DfsuFileType.Dfsu3DSigma || dfsuFile.DfsuFileType == DfsuFileType.Dfsu3DSigmaZ)
            {
                //reportTag.Error = TaskRunnerServiceRes.MIKE3NotImplementedYet;


                // doing type 32
                var coordList = (from el in ElementList
                                 where el.Type == 32
                                 orderby el.NodeList[0].Z
                                 select new { X1 = el.NodeList[0].X, X2 = el.NodeList[1].X, X3 = el.NodeList[2].X }).Distinct().ToList();

                foreach (var coord in coordList)
                {
                    int Layer = 1;
                    List<Element> ColumnElementList = (from el1 in ElementList
                                                       where el1.Type == 32
                                                       && el1.NodeList[0].X == coord.X1
                                                       && el1.NodeList[1].X == coord.X2
                                                       && el1.NodeList[2].X == coord.X3
                                                       orderby dfsuFile.Z[el1.NodeList[0].ID - 1] descending
                                                       select el1).ToList<Element>();

                    for (int j = 0; j < ColumnElementList.Count; j++)
                    {
                        ColumnElementList[j].Flag = 1;

                        ElementLayer elementLayer = new ElementLayer();
                        elementLayer.Layer = Layer;
                        elementLayer.ZMin = (from nz in ColumnElementList[j].NodeList select dfsuFile.Z[nz.ID - 1]).Min();
                        elementLayer.ZMax = (from nz in ColumnElementList[j].NodeList select dfsuFile.Z[nz.ID - 1]).Max();
                        elementLayer.Element = ColumnElementList[j];
                        ElementLayerList.Add(elementLayer);

                        NodeLayer nl3 = new NodeLayer();
                        nl3.Layer = Layer;
                        nl3.Z = dfsuFile.Z[ColumnElementList[j].NodeList[3].ID - 1];
                        nl3.Node = ColumnElementList[j].NodeList[3];

                        NodeLayer nl4 = new NodeLayer();
                        nl4.Layer = Layer;
                        nl4.Z = dfsuFile.Z[ColumnElementList[j].NodeList[4].ID - 1];
                        nl4.Node = ColumnElementList[j].NodeList[4];

                        NodeLayer nl5 = new NodeLayer();
                        nl5.Layer = Layer;
                        nl5.Z = dfsuFile.Z[ColumnElementList[j].NodeList[5].ID - 1];
                        nl5.Node = ColumnElementList[j].NodeList[5];

                        topNodeLayerList.Add(nl3);
                        topNodeLayerList.Add(nl4);
                        topNodeLayerList.Add(nl5);

                        NodeLayer nl0 = new NodeLayer();
                        nl0.Layer = Layer;
                        nl0.Z = dfsuFile.Z[ColumnElementList[j].NodeList[0].ID - 1];
                        nl0.Node = ColumnElementList[j].NodeList[0];

                        NodeLayer nl1 = new NodeLayer();
                        nl1.Layer = Layer;
                        nl1.Z = dfsuFile.Z[ColumnElementList[j].NodeList[1].ID - 1];
                        nl1.Node = ColumnElementList[j].NodeList[1];

                        NodeLayer nl2 = new NodeLayer();
                        nl2.Layer = Layer;
                        nl2.Z = dfsuFile.Z[ColumnElementList[j].NodeList[2].ID - 1];
                        nl2.Node = ColumnElementList[j].NodeList[2];

                        bottomNodeLayerList.Add(nl0);
                        bottomNodeLayerList.Add(nl1);
                        bottomNodeLayerList.Add(nl2);

                        Layer += 1;
                    }

                }

                var coordList2 = (from el in ElementList
                                  where el.Type == 33
                                  orderby el.NodeList[0].Z
                                  select new { X1 = el.NodeList[0].X, X2 = el.NodeList[1].X, X3 = el.NodeList[2].X, X4 = el.NodeList[3].X }).Distinct().ToList();

                foreach (var coord in coordList2)
                {
                    int Layer = 1;
                    List<Element> ColumElementList = (from el1 in ElementList
                                                      where el1.Type == 33
                                                      && el1.NodeList[0].X == coord.X1
                                                      && el1.NodeList[1].X == coord.X2
                                                      && el1.NodeList[2].X == coord.X3
                                                      && el1.NodeList[3].X == coord.X4
                                                      orderby dfsuFile.Z[el1.NodeList[0].ID - 1] descending
                                                      select el1).ToList<Element>();

                    for (int j = 0; j < ColumElementList.Count; j++)
                    {
                        ElementLayer elementLayer = new ElementLayer();
                        elementLayer.Layer = Layer;
                        elementLayer.ZMin = (from nz in ColumElementList[j].NodeList select dfsuFile.Z[nz.ID - 1]).Min();
                        elementLayer.ZMax = (from nz in ColumElementList[j].NodeList select dfsuFile.Z[nz.ID - 1]).Max();
                        elementLayer.Element = ColumElementList[j];
                        ElementLayerList.Add(elementLayer);

                        NodeLayer nl4 = new NodeLayer();
                        nl4.Layer = Layer;
                        nl4.Z = 0;
                        nl4.Node = ColumElementList[j].NodeList[4];

                        NodeLayer nl5 = new NodeLayer();
                        nl5.Layer = Layer;
                        nl5.Z = 0;
                        nl5.Node = ColumElementList[j].NodeList[5];

                        NodeLayer nl6 = new NodeLayer();
                        nl6.Layer = Layer;
                        nl6.Z = 0;
                        nl6.Node = ColumElementList[j].NodeList[6];

                        NodeLayer nl7 = new NodeLayer();
                        nl7.Layer = Layer;
                        nl7.Z = 0;
                        nl7.Node = ColumElementList[j].NodeList[7];


                        topNodeLayerList.Add(nl4);
                        topNodeLayerList.Add(nl5);
                        topNodeLayerList.Add(nl6);
                        topNodeLayerList.Add(nl7);

                        NodeLayer nl0 = new NodeLayer();
                        nl0.Layer = Layer;
                        nl0.Z = dfsuFile.Z[ColumElementList[j].NodeList[0].ID - 1];
                        nl0.Node = ColumElementList[j].NodeList[0];

                        NodeLayer nl1 = new NodeLayer();
                        nl1.Layer = Layer;
                        nl1.Z = dfsuFile.Z[ColumElementList[j].NodeList[1].ID - 1];
                        nl1.Node = ColumElementList[j].NodeList[1];

                        NodeLayer nl2 = new NodeLayer();
                        nl2.Layer = Layer;
                        nl2.Z = dfsuFile.Z[ColumElementList[j].NodeList[2].ID - 1];
                        nl2.Node = ColumElementList[j].NodeList[2];

                        NodeLayer nl3 = new NodeLayer();
                        nl3.Layer = Layer;
                        nl3.Z = dfsuFile.Z[ColumElementList[j].NodeList[3].ID - 1];
                        nl3.Node = ColumElementList[j].NodeList[3];

                        bottomNodeLayerList.Add(nl0);
                        bottomNodeLayerList.Add(nl1);
                        bottomNodeLayerList.Add(nl2);
                        bottomNodeLayerList.Add(nl3);

                        Layer += 1;
                    }

                }


                List<ElementLayer> TempElementLayerList = (from el in ElementLayerList
                                                           orderby el.Element.ID
                                                           select el).Distinct().ToList();

                //ElementLayerList = new List<ElementLayer>();
                int OldElemID = 0;
                foreach (ElementLayer el in TempElementLayerList)
                {
                    if (OldElemID == el.Element.ID)
                    {
                        ElementLayerList.Remove(el);
                    }
                    OldElemID = el.Element.ID;
                }

                List<NodeLayer> TempNodeLayerList = (from nl in topNodeLayerList
                                                     orderby nl.Node.ID
                                                     select nl).Distinct().ToList();

                topNodeLayerList = new List<NodeLayer>();
                int OldID = 0;
                foreach (NodeLayer nl in TempNodeLayerList)
                {
                    if (OldID != nl.Node.ID)
                    {
                        topNodeLayerList.Add(nl);
                        OldID = nl.Node.ID;
                    }
                }

                if (bottomNodeLayerList.Count() > 0)
                {
                    TempNodeLayerList = (from nl in bottomNodeLayerList
                                         orderby nl.Node.ID
                                         select nl).Distinct().ToList();

                    bottomNodeLayerList = new List<NodeLayer>();
                    OldID = 0;
                    foreach (NodeLayer nl in TempNodeLayerList)
                    {
                        if (OldID != nl.Node.ID)
                        {
                            bottomNodeLayerList.Add(nl);
                            OldID = nl.Node.ID;
                        }
                    }
                }

            }
            else if (dfsuFile.DfsuFileType == DfsuFileType.Dfsu2D)
            {
                int Layer = 1;
                for (int j = 0; j < ElementList.Count; j++)
                {
                    ElementLayer elementLayer = new ElementLayer();
                    elementLayer.Layer = Layer;
                    elementLayer.ZMin = 0;
                    elementLayer.ZMax = 0;
                    elementLayer.Element = ElementList[j];
                    ElementLayerList.Add(elementLayer);

                    // doing Nodes
                    if (ElementList[j].Type == 21)
                    {
                        NodeLayer nl0 = new NodeLayer();
                        nl0.Layer = Layer;
                        nl0.Z = 0;
                        nl0.Node = ElementList[j].NodeList[0];

                        NodeLayer nl1 = new NodeLayer();
                        nl1.Layer = Layer;
                        nl1.Z = 0;
                        nl1.Node = ElementList[j].NodeList[1];

                        NodeLayer nl2 = new NodeLayer();
                        nl2.Layer = Layer;
                        nl2.Z = 0;
                        nl2.Node = ElementList[j].NodeList[2];

                        topNodeLayerList.Add(nl0);
                        topNodeLayerList.Add(nl1);
                        topNodeLayerList.Add(nl2);
                    }
                    else if (ElementList[j].Type == 24)
                    {
                        NotUsed = string.Format(TaskRunnerServiceRes.NotImplementedYet, dfsuFile.NumberOfSigmaLayers.ToString());
                        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("NotImplementedYet", dfsuFile.NumberOfSigmaLayers.ToString());
                        return false;
                    }
                    else if (ElementList[j].Type == 25)
                    {
                        NodeLayer nl0 = new NodeLayer();
                        nl0.Layer = Layer;
                        nl0.Z = 0;
                        nl0.Node = ElementList[j].NodeList[0];

                        NodeLayer nl1 = new NodeLayer();
                        nl1.Layer = Layer;
                        nl1.Z = 0;
                        nl1.Node = ElementList[j].NodeList[1];

                        NodeLayer nl2 = new NodeLayer();
                        nl2.Layer = Layer;
                        nl2.Z = 0;
                        nl2.Node = ElementList[j].NodeList[2];

                        NodeLayer nl3 = new NodeLayer();
                        nl3.Layer = Layer;
                        nl3.Z = 0;
                        nl3.Node = ElementList[j].NodeList[3];


                        topNodeLayerList.Add(nl0);
                        topNodeLayerList.Add(nl1);
                        topNodeLayerList.Add(nl2);
                        topNodeLayerList.Add(nl3);
                    }
                }

                List<ElementLayer> TempElementLayerList = (from el in ElementLayerList
                                                           orderby el.Element.ID
                                                           select el).Distinct().ToList();

                ElementLayerList = new List<ElementLayer>();
                foreach (ElementLayer el in TempElementLayerList)
                {
                    ElementLayerList.Add(el);
                }

                List<NodeLayer> TempNodeLayerList = (from nl in topNodeLayerList
                                                     orderby nl.Node.ID
                                                     select nl).Distinct().ToList();

                topNodeLayerList = new List<NodeLayer>();
                int OldID = 0;
                foreach (NodeLayer nl in TempNodeLayerList)
                {
                    if (OldID != nl.Node.ID)
                    {
                        topNodeLayerList.Add(nl);
                        OldID = nl.Node.ID;
                    }
                }
            }
            else
            {
                NotUsed = string.Format(TaskRunnerServiceRes.NotImplementedYet, dfsuFile.NumberOfSigmaLayers.ToString());
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("NotImplementedYet", dfsuFile.NumberOfSigmaLayers.ToString());
                return false;
            }

            return true;
        }
        public bool FillElementListAndNodeList(DfsuFile dfsuFile, List<Element> ElementList, List<Node> NodeList)
        {
            for (int i = 0; i < dfsuFile.NumberOfNodes; i++)
            {
                Node n = new Node()
                {
                    Code = dfsuFile.Code[i],
                    ID = dfsuFile.NodeIds[i],
                    X = (float)dfsuFile.X[i],
                    Y = (float)dfsuFile.Y[i],
                    Z = dfsuFile.Z[i],
                    Value = 0,
                    ConnectNodeList = new List<Node>(),
                    ElementList = new List<Element>()
                };
                NodeList.Add(n);
            }

            for (int i = 0; i < dfsuFile.NumberOfElements; i++)
            {
                Element el = new Element()
                {
                    ID = dfsuFile.ElementIds[i],
                    Type = dfsuFile.ElementType[i],
                    Value = 0,
                    NodeList = new List<Node>(),
                    NumbOfNodes = 0
                };
                ElementList.Add(el);
            }

            for (int i = 0; i < dfsuFile.NumberOfElements; i++)
            {
                int CountNode = 0;
                for (int j = 0; j < dfsuFile.ElementTable[i].Count(); j++)
                {
                    CountNode += 1;
                    ElementList[i].NodeList.Add(NodeList[dfsuFile.ElementTable[i][j] - 1]);
                    if (!NodeList[dfsuFile.ElementTable[i][j] - 1].ElementList.Contains(ElementList[i]))
                    {
                        NodeList[dfsuFile.ElementTable[i][j] - 1].ElementList.Add(ElementList[i]);
                    }
                    for (int k = 0; k < dfsuFile.ElementTable[i].Count(); k++)
                    {
                        if (k != j)
                        {
                            if (!NodeList[dfsuFile.ElementTable[i][j] - 1].ConnectNodeList.Contains(NodeList[dfsuFile.ElementTable[i][k] - 1]))
                            {
                                NodeList[dfsuFile.ElementTable[i][j] - 1].ConnectNodeList.Add(NodeList[dfsuFile.ElementTable[i][k] - 1]);
                            }
                        }
                    }
                }
                ElementList[i].NumbOfNodes = CountNode;
            }

            return true;
        }
        public bool FillRequiredList(DfsuFile dfsuFile, List<ElementLayer> elementLayerList, List<Element> elementList, List<Node> nodeList, List<NodeLayer> topNodeLayerList, List<NodeLayer> bottomNodeLayerList)
        {
            if (!FillElementListAndNodeList(dfsuFile, elementList, nodeList))
            {
                return false;
            }
            if (!FillElementLayerList(dfsuFile, elementList, elementLayerList, topNodeLayerList, bottomNodeLayerList))
            {
                return false;
            }

            return true;
        }
        private bool FillVectors21_32(Element el, List<Element> UniqueElementList, float ContourValue, bool Is3D, bool IsTop)
        {
            string NotUsed = "";

            Node Node0 = new Node();
            Node Node1 = new Node();
            Node Node2 = new Node();
            if (Is3D && IsTop)
            {
                Node0 = el.NodeList[3];
                Node1 = el.NodeList[4];
                Node2 = el.NodeList[5];
            }
            else
            {
                Node0 = el.NodeList[0];
                Node1 = el.NodeList[1];
                Node2 = el.NodeList[2];
            }

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

            if (Node0.Value >= ContourValue && Node1.Value >= ContourValue && Node2.Value >= ContourValue)
            {
                if (Node0.Code != 0 && Node1.Code != 0 && ElemCount01 == 1)
                {
                    ForwardVector.Add(Node0.ID.ToString() + "," + Node1.ID.ToString(), new Vector() { StartNode = Node0, EndNode = Node1 });
                    BackwardVector.Add(Node1.ID.ToString() + "," + Node0.ID.ToString(), new Vector() { StartNode = Node1, EndNode = Node0 });
                }
                if (Node0.Code != 0 && Node2.Code != 0 && ElemCount02 == 1)
                {
                    ForwardVector.Add(Node0.ID.ToString() + "," + Node2.ID.ToString(), new Vector() { StartNode = Node0, EndNode = Node2 });
                    BackwardVector.Add(Node2.ID.ToString() + "," + Node0.ID.ToString(), new Vector() { StartNode = Node2, EndNode = Node0 });
                }
                if (Node1.Code != 0 && Node2.Code != 0 && ElemCount12 == 1)
                {
                    ForwardVector.Add(Node1.ID.ToString() + "," + Node2.ID.ToString(), new Vector() { StartNode = Node1, EndNode = Node2 });
                    BackwardVector.Add(Node2.ID.ToString() + "," + Node1.ID.ToString(), new Vector() { StartNode = Node2, EndNode = Node1 });
                }
            }
            else if (Node0.Value >= ContourValue && Node1.Value >= ContourValue && Node2.Value < ContourValue)
            {
                if (Node0.Code != 0 && Node1.Code != 0 && ElemCount01 == 1)
                {
                    ForwardVector.Add(Node0.ID.ToString() + "," + Node1.ID.ToString(), new Vector() { StartNode = Node0, EndNode = Node1 });
                    BackwardVector.Add(Node1.ID.ToString() + "," + Node0.ID.ToString(), new Vector() { StartNode = Node1, EndNode = Node0 });
                }
                Node TempInt1 = InterpolatedContourNodeList.Where(IntNode => IntNode.ID == Node0.ID * 1000000 + Node2.ID).First();
                if (Node0.Code != 0 && Node2.Code != 0 && ElemCount02 == 1)
                {
                    if (TempInt1 != null)
                    {
                        ForwardVector.Add(Node0.ID.ToString() + "," + TempInt1.ID.ToString(), new Vector() { StartNode = Node0, EndNode = TempInt1 });
                        BackwardVector.Add(TempInt1.ID.ToString() + "," + Node0.ID.ToString(), new Vector() { StartNode = TempInt1, EndNode = Node0 });
                    }


                }
                Node TempInt2 = InterpolatedContourNodeList.Where(IntNode => IntNode.ID == Node1.ID * 1000000 + Node2.ID).First();
                if (Node1.Code != 0 && Node2.Code != 0 && ElemCount12 == 1)
                {
                    if (TempInt2 != null)
                    {
                        ForwardVector.Add(Node1.ID.ToString() + "," + TempInt2.ID.ToString(), new Vector() { StartNode = Node1, EndNode = TempInt2 });
                        BackwardVector.Add(TempInt2.ID.ToString() + "," + Node1.ID.ToString(), new Vector() { StartNode = TempInt2, EndNode = Node1 });
                    }
                }

                if (TempInt1 != null && TempInt2 != null)
                {
                    ForwardVector.Add(TempInt1.ID.ToString() + "," + TempInt2.ID.ToString(), new Vector() { StartNode = TempInt1, EndNode = TempInt2 });
                    BackwardVector.Add(TempInt2.ID.ToString() + "," + TempInt1.ID.ToString(), new Vector() { StartNode = TempInt2, EndNode = TempInt1 });
                }
            }
            else if (Node0.Value >= ContourValue && Node1.Value < ContourValue && Node2.Value >= ContourValue)
            {
                if (Node0.Code != 0 && Node2.Code != 0 && ElemCount02 == 1)
                {
                    ForwardVector.Add(Node0.ID.ToString() + "," + Node2.ID.ToString(), new Vector() { StartNode = Node0, EndNode = Node2 });
                    BackwardVector.Add(Node2.ID.ToString() + "," + Node0.ID.ToString(), new Vector() { StartNode = Node2, EndNode = Node0 });
                }
                Node TempInt1 = InterpolatedContourNodeList.Where(IntNode => IntNode.ID == Node0.ID * 1000000 + Node1.ID).First();
                if (Node0.Code != 0 && Node1.Code != 0 && ElemCount01 == 1)
                {
                    if (TempInt1 != null)
                    {
                        ForwardVector.Add(Node0.ID.ToString() + "," + TempInt1.ID.ToString(), new Vector() { StartNode = Node0, EndNode = TempInt1 });
                        BackwardVector.Add(TempInt1.ID.ToString() + "," + Node0.ID.ToString(), new Vector() { StartNode = TempInt1, EndNode = Node0 });
                    }


                }
                Node TempInt2 = InterpolatedContourNodeList.Where(IntNode => IntNode.ID == Node2.ID * 1000000 + Node1.ID).First();
                if (Node1.Code != 0 && Node2.Code != 0 && ElemCount12 == 1)
                {
                    if (TempInt2 != null)
                    {
                        ForwardVector.Add(Node2.ID.ToString() + "," + TempInt2.ID.ToString(), new Vector() { StartNode = Node2, EndNode = TempInt2 });
                        BackwardVector.Add(TempInt2.ID.ToString() + "," + Node2.ID.ToString(), new Vector() { StartNode = TempInt2, EndNode = Node2 });
                    }
                }

                if (TempInt1 != null && TempInt2 != null)
                {
                    ForwardVector.Add(TempInt1.ID.ToString() + "," + TempInt2.ID.ToString(), new Vector() { StartNode = TempInt1, EndNode = TempInt2 });
                    BackwardVector.Add(TempInt2.ID.ToString() + "," + TempInt1.ID.ToString(), new Vector() { StartNode = TempInt2, EndNode = TempInt1 });
                }
            }
            else if (Node0.Value >= ContourValue && Node1.Value < ContourValue && Node2.Value < ContourValue)
            {
                Node TempInt1 = InterpolatedContourNodeList.Where(IntNode => IntNode.ID == Node0.ID * 1000000 + Node1.ID).First();
                if (Node0.Code != 0 && Node1.Code != 0 && ElemCount01 == 1)
                {
                    if (TempInt1 != null)
                    {
                        ForwardVector.Add(Node0.ID.ToString() + "," + TempInt1.ID.ToString(), new Vector() { StartNode = Node0, EndNode = TempInt1 });
                        BackwardVector.Add(TempInt1.ID.ToString() + "," + Node0.ID.ToString(), new Vector() { StartNode = TempInt1, EndNode = Node0 });
                    }


                }
                Node TempInt2 = InterpolatedContourNodeList.Where(IntNode => IntNode.ID == Node0.ID * 1000000 + Node2.ID).First();
                if (Node0.Code != 0 && Node2.Code != 0 && ElemCount02 == 1)
                {
                    if (TempInt2 != null)
                    {
                        ForwardVector.Add(Node0.ID.ToString() + "," + TempInt2.ID.ToString(), new Vector() { StartNode = Node0, EndNode = TempInt2 });
                        BackwardVector.Add(TempInt2.ID.ToString() + "," + Node0.ID.ToString(), new Vector() { StartNode = TempInt2, EndNode = Node0 });
                    }
                }

                if (TempInt1 != null && TempInt2 != null)
                {
                    ForwardVector.Add(TempInt1.ID.ToString() + "," + TempInt2.ID.ToString(), new Vector() { StartNode = TempInt1, EndNode = TempInt2 });
                    BackwardVector.Add(TempInt2.ID.ToString() + "," + TempInt1.ID.ToString(), new Vector() { StartNode = TempInt2, EndNode = TempInt1 });
                }
            }
            else if (Node0.Value < ContourValue && Node1.Value >= ContourValue && Node2.Value >= ContourValue)
            {
                if (Node1.Code != 0 && Node2.Code != 0 && ElemCount12 == 1)
                {
                    ForwardVector.Add(Node1.ID.ToString() + "," + Node2.ID.ToString(), new Vector() { StartNode = Node1, EndNode = Node2 });
                    BackwardVector.Add(Node2.ID.ToString() + "," + Node1.ID.ToString(), new Vector() { StartNode = Node2, EndNode = Node1 });
                }
                Node TempInt1 = InterpolatedContourNodeList.Where(IntNode => IntNode.ID == Node1.ID * 1000000 + Node0.ID).First();
                if (Node0.Code != 0 && Node1.Code != 0 && ElemCount01 == 1)
                {
                    if (TempInt1 != null)
                    {
                        ForwardVector.Add(Node1.ID.ToString() + "," + TempInt1.ID.ToString(), new Vector() { StartNode = Node1, EndNode = TempInt1 });
                        BackwardVector.Add(TempInt1.ID.ToString() + "," + Node1.ID.ToString(), new Vector() { StartNode = TempInt1, EndNode = Node1 });
                    }


                }
                Node TempInt2 = InterpolatedContourNodeList.Where(IntNode => IntNode.ID == Node2.ID * 1000000 + Node0.ID).First();
                if (Node0.Code != 0 && Node2.Code != 0 && ElemCount02 == 1)
                {
                    if (TempInt2 != null)
                    {
                        ForwardVector.Add(Node2.ID.ToString() + "," + TempInt2.ID.ToString(), new Vector() { StartNode = Node2, EndNode = TempInt2 });
                        BackwardVector.Add(TempInt2.ID.ToString() + "," + Node2.ID.ToString(), new Vector() { StartNode = TempInt2, EndNode = Node2 });
                    }
                }

                if (TempInt1 != null && TempInt2 != null)
                {
                    ForwardVector.Add(TempInt1.ID.ToString() + "," + TempInt2.ID.ToString(), new Vector() { StartNode = TempInt1, EndNode = TempInt2 });
                    BackwardVector.Add(TempInt2.ID.ToString() + "," + TempInt1.ID.ToString(), new Vector() { StartNode = TempInt2, EndNode = TempInt1 });
                }
            }
            else if (Node0.Value < ContourValue && Node1.Value >= ContourValue && Node2.Value < ContourValue)
            {
                Node TempInt1 = InterpolatedContourNodeList.Where(IntNode => IntNode.ID == Node1.ID * 1000000 + Node0.ID).First();
                if (Node0.Code != 0 && Node1.Code != 0 && ElemCount01 == 1)
                {
                    if (TempInt1 != null)
                    {
                        ForwardVector.Add(Node1.ID.ToString() + "," + TempInt1.ID.ToString(), new Vector() { StartNode = Node1, EndNode = TempInt1 });
                        BackwardVector.Add(TempInt1.ID.ToString() + "," + Node1.ID.ToString(), new Vector() { StartNode = TempInt1, EndNode = Node1 });
                    }


                }
                Node TempInt2 = InterpolatedContourNodeList.Where(IntNode => IntNode.ID == Node1.ID * 1000000 + Node2.ID).First();
                if (Node1.Code != 0 && Node2.Code != 0 && ElemCount12 == 1)
                {
                    if (TempInt2 != null)
                    {
                        ForwardVector.Add(Node1.ID.ToString() + "," + TempInt2.ID.ToString(), new Vector() { StartNode = Node1, EndNode = TempInt2 });
                        BackwardVector.Add(TempInt2.ID.ToString() + "," + Node1.ID.ToString(), new Vector() { StartNode = TempInt2, EndNode = Node1 });
                    }
                }

                if (TempInt1 != null && TempInt2 != null)
                {
                    ForwardVector.Add(TempInt1.ID.ToString() + "," + TempInt2.ID.ToString(), new Vector() { StartNode = TempInt1, EndNode = TempInt2 });
                    BackwardVector.Add(TempInt2.ID.ToString() + "," + TempInt1.ID.ToString(), new Vector() { StartNode = TempInt2, EndNode = TempInt1 });
                }
            }
            else if (Node0.Value < ContourValue && Node1.Value < ContourValue && Node2.Value >= ContourValue)
            {
                Node TempInt1 = InterpolatedContourNodeList.Where(IntNode => IntNode.ID == Node2.ID * 1000000 + Node0.ID).First();
                if (Node0.Code != 0 && Node2.Code != 0 && ElemCount02 == 1)
                {
                    if (TempInt1 != null)
                    {
                        ForwardVector.Add(Node2.ID.ToString() + "," + TempInt1.ID.ToString(), new Vector() { StartNode = Node2, EndNode = TempInt1 });
                        BackwardVector.Add(TempInt1.ID.ToString() + "," + Node2.ID.ToString(), new Vector() { StartNode = TempInt1, EndNode = Node2 });
                    }


                }
                Node TempInt2 = InterpolatedContourNodeList.Where(IntNode => IntNode.ID == Node2.ID * 1000000 + Node1.ID).First();
                if (Node1.Code != 0 && Node2.Code != 0 && ElemCount12 == 1)
                {
                    if (TempInt2 != null)
                    {
                        ForwardVector.Add(Node2.ID.ToString() + "," + TempInt2.ID.ToString(), new Vector() { StartNode = Node2, EndNode = TempInt2 });
                        BackwardVector.Add(TempInt2.ID.ToString() + "," + Node2.ID.ToString(), new Vector() { StartNode = TempInt2, EndNode = Node2 });
                    }
                }

                if (TempInt1 != null && TempInt2 != null)
                {
                    ForwardVector.Add(TempInt1.ID.ToString() + "," + TempInt2.ID.ToString(), new Vector() { StartNode = TempInt1, EndNode = TempInt2 });
                    BackwardVector.Add(TempInt2.ID.ToString() + "," + TempInt1.ID.ToString(), new Vector() { StartNode = TempInt2, EndNode = TempInt1 });
                }
            }
            else if (Node0.Value < ContourValue && Node1.Value < ContourValue && Node2.Value < ContourValue)
            {
                // no vector to create
            }
            else
            {
                NotUsed = TaskRunnerServiceRes.AllNodesAreSmallerThanContourValue;
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageList("AllNodesAreSmallerThanContourValue");
                return false;
            }

            return true;
        }
        private bool FillVectors25_33(Element el, List<Element> UniqueElementList, float ContourValue, bool Is3D, bool IsTop)
        {
            string NotUsed = "";

            Node Node0 = new Node();
            Node Node1 = new Node();
            Node Node2 = new Node();
            Node Node3 = new Node();

            if (Is3D && IsTop)
            {
                Node0 = el.NodeList[4];
                Node1 = el.NodeList[5];
                Node2 = el.NodeList[6];
                Node3 = el.NodeList[7];
            }
            else
            {
                Node0 = el.NodeList[0];
                Node1 = el.NodeList[1];
                Node2 = el.NodeList[2];
                Node3 = el.NodeList[3];
            }

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

            if (Node0.Value >= ContourValue && Node1.Value >= ContourValue && Node2.Value >= ContourValue && Node3.Value >= ContourValue)
            {
                if (Node0.Code != 0 && Node1.Code != 0 && ElemCount01 == 1)
                {
                    ForwardVector.Add(Node0.ID.ToString() + "," + Node1.ID.ToString(), new Vector() { StartNode = Node0, EndNode = Node1 });
                    BackwardVector.Add(Node1.ID.ToString() + "," + Node0.ID.ToString(), new Vector() { StartNode = Node1, EndNode = Node0 });
                }
                if (Node0.Code != 0 && Node2.Code != 0 && ElemCount03 == 1)
                {
                    ForwardVector.Add(Node0.ID.ToString() + "," + Node3.ID.ToString(), new Vector() { StartNode = Node0, EndNode = Node3 });
                    BackwardVector.Add(Node3.ID.ToString() + "," + Node0.ID.ToString(), new Vector() { StartNode = Node3, EndNode = Node0 });
                }
                if (Node1.Code != 0 && Node2.Code != 0 && ElemCount12 == 1)
                {
                    ForwardVector.Add(Node1.ID.ToString() + "," + Node2.ID.ToString(), new Vector() { StartNode = Node1, EndNode = Node2 });
                    BackwardVector.Add(Node2.ID.ToString() + "," + Node1.ID.ToString(), new Vector() { StartNode = Node2, EndNode = Node1 });
                }
                if (Node2.Code != 0 && Node3.Code != 0 && ElemCount23 == 1)
                {
                    ForwardVector.Add(Node2.ID.ToString() + "," + Node3.ID.ToString(), new Vector() { StartNode = Node2, EndNode = Node3 });
                    BackwardVector.Add(Node3.ID.ToString() + "," + Node2.ID.ToString(), new Vector() { StartNode = Node3, EndNode = Node2 });
                }
            }
            else if (Node0.Value >= ContourValue && Node1.Value >= ContourValue && Node2.Value >= ContourValue && Node3.Value < ContourValue)
            {
                if (Node0.Code != 0 && Node1.Code != 0 && ElemCount01 == 1)
                {
                    ForwardVector.Add(Node0.ID.ToString() + "," + Node1.ID.ToString(), new Vector() { StartNode = Node0, EndNode = Node1 });
                    BackwardVector.Add(Node1.ID.ToString() + "," + Node0.ID.ToString(), new Vector() { StartNode = Node1, EndNode = Node0 });
                }
                if (Node1.Code != 0 && Node2.Code != 0 && ElemCount12 == 1)
                {
                    ForwardVector.Add(Node1.ID.ToString() + "," + Node2.ID.ToString(), new Vector() { StartNode = Node1, EndNode = Node2 });
                    BackwardVector.Add(Node2.ID.ToString() + "," + Node1.ID.ToString(), new Vector() { StartNode = Node2, EndNode = Node1 });
                }
                Node TempInt1 = InterpolatedContourNodeList.Where(IntNode => IntNode.ID == Node0.ID * 1000000 + Node3.ID).First();
                if (Node0.Code != 0 && Node3.Code != 0 && ElemCount03 == 1)
                {
                    if (TempInt1 != null)
                    {
                        ForwardVector.Add(Node0.ID.ToString() + "," + TempInt1.ID.ToString(), new Vector() { StartNode = Node0, EndNode = TempInt1 });
                        BackwardVector.Add(TempInt1.ID.ToString() + "," + Node0.ID.ToString(), new Vector() { StartNode = TempInt1, EndNode = Node0 });
                    }


                }
                Node TempInt2 = InterpolatedContourNodeList.Where(IntNode => IntNode.ID == Node2.ID * 1000000 + Node3.ID).First();
                if (Node2.Code != 0 && Node3.Code != 0 && ElemCount23 == 1)
                {
                    if (TempInt2 != null)
                    {
                        ForwardVector.Add(Node2.ID.ToString() + "," + TempInt2.ID.ToString(), new Vector() { StartNode = Node2, EndNode = TempInt2 });
                        BackwardVector.Add(TempInt2.ID.ToString() + "," + Node2.ID.ToString(), new Vector() { StartNode = TempInt2, EndNode = Node2 });
                    }
                }

                if (TempInt1 != null && TempInt2 != null)
                {
                    ForwardVector.Add(TempInt1.ID.ToString() + "," + TempInt2.ID.ToString(), new Vector() { StartNode = TempInt1, EndNode = TempInt2 });
                    BackwardVector.Add(TempInt2.ID.ToString() + "," + TempInt1.ID.ToString(), new Vector() { StartNode = TempInt2, EndNode = TempInt1 });
                }
            }
            else if (Node0.Value >= ContourValue && Node1.Value >= ContourValue && Node2.Value < ContourValue && Node3.Value >= ContourValue)
            {
                if (Node0.Code != 0 && Node1.Code != 0 && ElemCount01 == 1)
                {
                    ForwardVector.Add(Node0.ID.ToString() + "," + Node1.ID.ToString(), new Vector() { StartNode = Node0, EndNode = Node1 });
                    BackwardVector.Add(Node1.ID.ToString() + "," + Node0.ID.ToString(), new Vector() { StartNode = Node1, EndNode = Node0 });
                }
                if (Node0.Code != 0 && Node3.Code != 0 && ElemCount03 == 1)
                {
                    ForwardVector.Add(Node0.ID.ToString() + "," + Node3.ID.ToString(), new Vector() { StartNode = Node0, EndNode = Node3 });
                    BackwardVector.Add(Node3.ID.ToString() + "," + Node0.ID.ToString(), new Vector() { StartNode = Node3, EndNode = Node0 });
                }
                Node TempInt1 = InterpolatedContourNodeList.Where(IntNode => IntNode.ID == Node1.ID * 1000000 + Node2.ID).First();
                if (Node1.Code != 0 && Node2.Code != 0 && ElemCount12 == 1)
                {
                    if (TempInt1 != null)
                    {
                        ForwardVector.Add(Node1.ID.ToString() + "," + TempInt1.ID.ToString(), new Vector() { StartNode = Node1, EndNode = TempInt1 });
                        BackwardVector.Add(TempInt1.ID.ToString() + "," + Node1.ID.ToString(), new Vector() { StartNode = TempInt1, EndNode = Node1 });
                    }


                }
                Node TempInt2 = InterpolatedContourNodeList.Where(IntNode => IntNode.ID == Node3.ID * 1000000 + Node2.ID).First();
                if (Node3.Code != 0 && Node2.Code != 0 && ElemCount23 == 1)
                {
                    if (TempInt2 != null)
                    {
                        ForwardVector.Add(Node3.ID.ToString() + "," + TempInt2.ID.ToString(), new Vector() { StartNode = Node3, EndNode = TempInt2 });
                        BackwardVector.Add(TempInt2.ID.ToString() + "," + Node3.ID.ToString(), new Vector() { StartNode = TempInt2, EndNode = Node3 });
                    }
                }

                if (TempInt1 != null && TempInt2 != null)
                {
                    ForwardVector.Add(TempInt1.ID.ToString() + "," + TempInt2.ID.ToString(), new Vector() { StartNode = TempInt1, EndNode = TempInt2 });
                    BackwardVector.Add(TempInt2.ID.ToString() + "," + TempInt1.ID.ToString(), new Vector() { StartNode = TempInt2, EndNode = TempInt1 });
                }
            }
            else if (Node0.Value >= ContourValue && Node1.Value < ContourValue && Node2.Value >= ContourValue && Node3.Value >= ContourValue)
            {
                if (Node0.Code != 0 && Node3.Code != 0 && ElemCount03 == 1)
                {
                    ForwardVector.Add(Node0.ID.ToString() + "," + Node3.ID.ToString(), new Vector() { StartNode = Node0, EndNode = Node3 });
                    BackwardVector.Add(Node3.ID.ToString() + "," + Node0.ID.ToString(), new Vector() { StartNode = Node3, EndNode = Node0 });
                }
                if (Node2.Code != 0 && Node3.Code != 0 && ElemCount23 == 1)
                {
                    ForwardVector.Add(Node2.ID.ToString() + "," + Node3.ID.ToString(), new Vector() { StartNode = Node2, EndNode = Node3 });
                    BackwardVector.Add(Node3.ID.ToString() + "," + Node2.ID.ToString(), new Vector() { StartNode = Node3, EndNode = Node2 });
                }
                Node TempInt1 = InterpolatedContourNodeList.Where(IntNode => IntNode.ID == Node0.ID * 1000000 + Node1.ID).First();
                if (Node0.Code != 0 && Node1.Code != 0 && ElemCount01 == 1)
                {
                    if (TempInt1 != null)
                    {
                        ForwardVector.Add(Node0.ID.ToString() + "," + TempInt1.ID.ToString(), new Vector() { StartNode = Node0, EndNode = TempInt1 });
                        BackwardVector.Add(TempInt1.ID.ToString() + "," + Node0.ID.ToString(), new Vector() { StartNode = TempInt1, EndNode = Node0 });
                    }


                }
                Node TempInt2 = InterpolatedContourNodeList.Where(IntNode => IntNode.ID == Node2.ID * 1000000 + Node1.ID).First();
                if (Node2.Code != 0 && Node1.Code != 0 && ElemCount12 == 1)
                {
                    if (TempInt2 != null)
                    {
                        ForwardVector.Add(Node2.ID.ToString() + "," + TempInt2.ID.ToString(), new Vector() { StartNode = Node2, EndNode = TempInt2 });
                        BackwardVector.Add(TempInt2.ID.ToString() + "," + Node2.ID.ToString(), new Vector() { StartNode = TempInt2, EndNode = Node2 });
                    }
                }

                if (TempInt1 != null && TempInt2 != null)
                {
                    ForwardVector.Add(TempInt1.ID.ToString() + "," + TempInt2.ID.ToString(), new Vector() { StartNode = TempInt1, EndNode = TempInt2 });
                    BackwardVector.Add(TempInt2.ID.ToString() + "," + TempInt1.ID.ToString(), new Vector() { StartNode = TempInt2, EndNode = TempInt1 });
                }
            }
            else if (Node0.Value < ContourValue && Node1.Value >= ContourValue && Node2.Value >= ContourValue && Node3.Value >= ContourValue)
            {
                if (Node1.Code != 0 && Node2.Code != 0 && ElemCount12 == 1)
                {
                    ForwardVector.Add(Node1.ID.ToString() + "," + Node2.ID.ToString(), new Vector() { StartNode = Node1, EndNode = Node2 });
                    BackwardVector.Add(Node2.ID.ToString() + "," + Node1.ID.ToString(), new Vector() { StartNode = Node2, EndNode = Node1 });
                }
                if (Node2.Code != 0 && Node3.Code != 0 && ElemCount23 == 1)
                {
                    ForwardVector.Add(Node2.ID.ToString() + "," + Node3.ID.ToString(), new Vector() { StartNode = Node2, EndNode = Node3 });
                    BackwardVector.Add(Node3.ID.ToString() + "," + Node2.ID.ToString(), new Vector() { StartNode = Node3, EndNode = Node2 });
                }
                Node TempInt1 = InterpolatedContourNodeList.Where(IntNode => IntNode.ID == Node1.ID * 1000000 + Node0.ID).First();
                if (Node1.Code != 0 && Node0.Code != 0 && ElemCount01 == 1)
                {
                    if (TempInt1 != null)
                    {
                        ForwardVector.Add(Node1.ID.ToString() + "," + TempInt1.ID.ToString(), new Vector() { StartNode = Node1, EndNode = TempInt1 });
                        BackwardVector.Add(TempInt1.ID.ToString() + "," + Node1.ID.ToString(), new Vector() { StartNode = TempInt1, EndNode = Node1 });
                    }


                }
                Node TempInt2 = InterpolatedContourNodeList.Where(IntNode => IntNode.ID == Node3.ID * 1000000 + Node0.ID).First();
                if (Node3.Code != 0 && Node0.Code != 0 && ElemCount03 == 1)
                {
                    if (TempInt2 != null)
                    {
                        ForwardVector.Add(Node3.ID.ToString() + "," + TempInt2.ID.ToString(), new Vector() { StartNode = Node3, EndNode = TempInt2 });
                        BackwardVector.Add(TempInt2.ID.ToString() + "," + Node3.ID.ToString(), new Vector() { StartNode = TempInt2, EndNode = Node3 });
                    }
                }

                if (TempInt1 != null && TempInt2 != null)
                {
                    ForwardVector.Add(TempInt1.ID.ToString() + "," + TempInt2.ID.ToString(), new Vector() { StartNode = TempInt1, EndNode = TempInt2 });
                    BackwardVector.Add(TempInt2.ID.ToString() + "," + TempInt1.ID.ToString(), new Vector() { StartNode = TempInt2, EndNode = TempInt1 });
                }
            }
            else if (Node0.Value >= ContourValue && Node1.Value >= ContourValue && Node2.Value < ContourValue && Node3.Value < ContourValue)
            {
                if (Node0.Code != 0 && Node1.Code != 0 && ElemCount01 == 1)
                {
                    ForwardVector.Add(Node0.ID.ToString() + "," + Node1.ID.ToString(), new Vector() { StartNode = Node0, EndNode = Node1 });
                    BackwardVector.Add(Node1.ID.ToString() + "," + Node0.ID.ToString(), new Vector() { StartNode = Node1, EndNode = Node0 });
                }
                Node TempInt1 = InterpolatedContourNodeList.Where(IntNode => IntNode.ID == Node0.ID * 1000000 + Node3.ID).First();
                if (Node0.Code != 0 && Node3.Code != 0 && ElemCount03 == 1)
                {
                    if (TempInt1 != null)
                    {
                        ForwardVector.Add(Node0.ID.ToString() + "," + TempInt1.ID.ToString(), new Vector() { StartNode = Node0, EndNode = TempInt1 });
                        BackwardVector.Add(TempInt1.ID.ToString() + "," + Node0.ID.ToString(), new Vector() { StartNode = TempInt1, EndNode = Node0 });
                    }


                }
                Node TempInt2 = InterpolatedContourNodeList.Where(IntNode => IntNode.ID == Node1.ID * 1000000 + Node2.ID).First();
                if (Node1.Code != 0 && Node2.Code != 0 && ElemCount12 == 1)
                {
                    if (TempInt2 != null)
                    {
                        ForwardVector.Add(Node1.ID.ToString() + "," + TempInt2.ID.ToString(), new Vector() { StartNode = Node1, EndNode = TempInt2 });
                        BackwardVector.Add(TempInt2.ID.ToString() + "," + Node1.ID.ToString(), new Vector() { StartNode = TempInt2, EndNode = Node1 });
                    }
                }

                if (TempInt1 != null && TempInt2 != null)
                {
                    ForwardVector.Add(TempInt1.ID.ToString() + "," + TempInt2.ID.ToString(), new Vector() { StartNode = TempInt1, EndNode = TempInt2 });
                    BackwardVector.Add(TempInt2.ID.ToString() + "," + TempInt1.ID.ToString(), new Vector() { StartNode = TempInt2, EndNode = TempInt1 });
                }
            }
            else if (Node0.Value >= ContourValue && Node1.Value < ContourValue && Node2.Value >= ContourValue && Node3.Value < ContourValue)
            {
                Node TempInt1 = InterpolatedContourNodeList.Where(IntNode => IntNode.ID == Node0.ID * 1000000 + Node3.ID).First();
                if (Node0.Code != 0 && Node3.Code != 0 && ElemCount03 == 1)
                {
                    if (TempInt1 != null)
                    {
                        ForwardVector.Add(Node0.ID.ToString() + "," + TempInt1.ID.ToString(), new Vector() { StartNode = Node0, EndNode = TempInt1 });
                        BackwardVector.Add(TempInt1.ID.ToString() + "," + Node0.ID.ToString(), new Vector() { StartNode = TempInt1, EndNode = Node0 });
                    }


                }
                Node TempInt2 = InterpolatedContourNodeList.Where(IntNode => IntNode.ID == Node0.ID * 1000000 + Node1.ID).First();
                if (Node0.Code != 0 && Node1.Code != 0 && ElemCount01 == 1)
                {
                    if (TempInt2 != null)
                    {
                        ForwardVector.Add(Node0.ID.ToString() + "," + TempInt2.ID.ToString(), new Vector() { StartNode = Node0, EndNode = TempInt2 });
                        BackwardVector.Add(TempInt2.ID.ToString() + "," + Node0.ID.ToString(), new Vector() { StartNode = TempInt2, EndNode = Node0 });
                    }
                }

                if (TempInt1 != null && TempInt2 != null)
                {
                    ForwardVector.Add(TempInt1.ID.ToString() + "," + TempInt2.ID.ToString(), new Vector() { StartNode = TempInt1, EndNode = TempInt2 });
                    BackwardVector.Add(TempInt2.ID.ToString() + "," + TempInt1.ID.ToString(), new Vector() { StartNode = TempInt2, EndNode = TempInt1 });
                }

                Node TempInt3 = InterpolatedContourNodeList.Where(IntNode => IntNode.ID == Node2.ID * 1000000 + Node1.ID).First();
                if (Node2.Code != 0 && Node1.Code != 0 && ElemCount12 == 1)
                {
                    if (TempInt3 != null)
                    {
                        ForwardVector.Add(Node2.ID.ToString() + "," + TempInt3.ID.ToString(), new Vector() { StartNode = Node2, EndNode = TempInt3 });
                        BackwardVector.Add(TempInt3.ID.ToString() + "," + Node2.ID.ToString(), new Vector() { StartNode = TempInt3, EndNode = Node2 });
                    }


                }
                Node TempInt4 = InterpolatedContourNodeList.Where(IntNode => IntNode.ID == Node2.ID * 1000000 + Node3.ID).First();
                if (Node2.Code != 0 && Node3.Code != 0 && ElemCount23 == 1)
                {
                    if (TempInt4 != null)
                    {
                        ForwardVector.Add(Node2.ID.ToString() + "," + TempInt4.ID.ToString(), new Vector() { StartNode = Node2, EndNode = TempInt4 });
                        BackwardVector.Add(TempInt4.ID.ToString() + "," + Node2.ID.ToString(), new Vector() { StartNode = TempInt4, EndNode = Node2 });
                    }
                }

                if (TempInt3 != null && TempInt4 != null)
                {
                    ForwardVector.Add(TempInt3.ID.ToString() + "," + TempInt4.ID.ToString(), new Vector() { StartNode = TempInt3, EndNode = TempInt4 });
                    BackwardVector.Add(TempInt4.ID.ToString() + "," + TempInt3.ID.ToString(), new Vector() { StartNode = TempInt4, EndNode = TempInt3 });
                }

            }
            else if (Node0.Value < ContourValue && Node1.Value >= ContourValue && Node2.Value >= ContourValue && Node3.Value < ContourValue)
            {
                Node TempInt1 = InterpolatedContourNodeList.Where(IntNode => IntNode.ID == Node1.ID * 1000000 + Node0.ID).First();
                if (Node1.Code != 0 && Node0.Code != 0 && ElemCount01 == 1)
                {
                    if (TempInt1 != null)
                    {
                        ForwardVector.Add(Node1.ID.ToString() + "," + TempInt1.ID.ToString(), new Vector() { StartNode = Node1, EndNode = TempInt1 });
                        BackwardVector.Add(TempInt1.ID.ToString() + "," + Node1.ID.ToString(), new Vector() { StartNode = TempInt1, EndNode = Node1 });
                    }


                }
                Node TempInt2 = InterpolatedContourNodeList.Where(IntNode => IntNode.ID == Node2.ID * 1000000 + Node3.ID).First();
                if (Node2.Code != 0 && Node3.Code != 0 && ElemCount23 == 1)
                {
                    if (TempInt2 != null)
                    {
                        ForwardVector.Add(Node2.ID.ToString() + "," + TempInt2.ID.ToString(), new Vector() { StartNode = Node2, EndNode = TempInt2 });
                        BackwardVector.Add(TempInt2.ID.ToString() + "," + Node2.ID.ToString(), new Vector() { StartNode = TempInt2, EndNode = Node2 });
                    }
                }

                if (TempInt1 != null && TempInt2 != null)
                {
                    ForwardVector.Add(TempInt1.ID.ToString() + "," + TempInt2.ID.ToString(), new Vector() { StartNode = TempInt1, EndNode = TempInt2 });
                    BackwardVector.Add(TempInt2.ID.ToString() + "," + TempInt1.ID.ToString(), new Vector() { StartNode = TempInt2, EndNode = TempInt1 });
                }

            }
            else if (Node0.Value >= ContourValue && Node1.Value < ContourValue && Node2.Value < ContourValue && Node3.Value >= ContourValue)
            {
                Node TempInt1 = InterpolatedContourNodeList.Where(IntNode => IntNode.ID == Node0.ID * 1000000 + Node1.ID).First();
                if (Node0.Code != 0 && Node1.Code != 0 && ElemCount01 == 1)
                {
                    if (TempInt1 != null)
                    {
                        ForwardVector.Add(Node0.ID.ToString() + "," + TempInt1.ID.ToString(), new Vector() { StartNode = Node0, EndNode = TempInt1 });
                        BackwardVector.Add(TempInt1.ID.ToString() + "," + Node0.ID.ToString(), new Vector() { StartNode = TempInt1, EndNode = Node0 });
                    }


                }
                Node TempInt2 = InterpolatedContourNodeList.Where(IntNode => IntNode.ID == Node3.ID * 1000000 + Node2.ID).First();
                if (Node3.Code != 0 && Node2.Code != 0 && ElemCount23 == 1)
                {
                    if (TempInt2 != null)
                    {
                        ForwardVector.Add(Node3.ID.ToString() + "," + TempInt2.ID.ToString(), new Vector() { StartNode = Node3, EndNode = TempInt2 });
                        BackwardVector.Add(TempInt2.ID.ToString() + "," + Node3.ID.ToString(), new Vector() { StartNode = TempInt2, EndNode = Node3 });
                    }
                }

                if (TempInt1 != null && TempInt2 != null)
                {
                    ForwardVector.Add(TempInt1.ID.ToString() + "," + TempInt2.ID.ToString(), new Vector() { StartNode = TempInt1, EndNode = TempInt2 });
                    BackwardVector.Add(TempInt2.ID.ToString() + "," + TempInt1.ID.ToString(), new Vector() { StartNode = TempInt2, EndNode = TempInt1 });
                }

            }
            else if (Node0.Value < ContourValue && Node1.Value >= ContourValue && Node2.Value < ContourValue && Node3.Value >= ContourValue)
            {
                Node TempInt1 = InterpolatedContourNodeList.Where(IntNode => IntNode.ID == Node3.ID * 1000000 + Node0.ID).First();
                if (Node3.Code != 0 && Node0.Code != 0 && ElemCount03 == 1)
                {
                    if (TempInt1 != null)
                    {
                        ForwardVector.Add(Node3.ID.ToString() + "," + TempInt1.ID.ToString(), new Vector() { StartNode = Node3, EndNode = TempInt1 });
                        BackwardVector.Add(TempInt1.ID.ToString() + "," + Node3.ID.ToString(), new Vector() { StartNode = TempInt1, EndNode = Node3 });
                    }


                }
                Node TempInt2 = InterpolatedContourNodeList.Where(IntNode => IntNode.ID == Node3.ID * 1000000 + Node2.ID).First();
                if (Node3.Code != 0 && Node2.Code != 0 && ElemCount23 == 1)
                {
                    if (TempInt2 != null)
                    {
                        ForwardVector.Add(Node3.ID.ToString() + "," + TempInt2.ID.ToString(), new Vector() { StartNode = Node3, EndNode = TempInt2 });
                        BackwardVector.Add(TempInt2.ID.ToString() + "," + Node3.ID.ToString(), new Vector() { StartNode = TempInt2, EndNode = Node3 });
                    }
                }

                if (TempInt1 != null && TempInt2 != null)
                {
                    ForwardVector.Add(TempInt1.ID.ToString() + "," + TempInt2.ID.ToString(), new Vector() { StartNode = TempInt1, EndNode = TempInt2 });
                    BackwardVector.Add(TempInt2.ID.ToString() + "," + TempInt1.ID.ToString(), new Vector() { StartNode = TempInt2, EndNode = TempInt1 });
                }

                Node TempInt3 = InterpolatedContourNodeList.Where(IntNode => IntNode.ID == Node1.ID * 1000000 + Node0.ID).First();
                if (Node1.Code != 0 && Node0.Code != 0 && ElemCount01 == 1)
                {
                    if (TempInt3 != null)
                    {
                        ForwardVector.Add(Node1.ID.ToString() + "," + TempInt3.ID.ToString(), new Vector() { StartNode = Node1, EndNode = TempInt3 });
                        BackwardVector.Add(TempInt3.ID.ToString() + "," + Node1.ID.ToString(), new Vector() { StartNode = TempInt3, EndNode = Node1 });
                    }


                }
                Node TempInt4 = InterpolatedContourNodeList.Where(IntNode => IntNode.ID == Node1.ID * 1000000 + Node2.ID).First();
                if (Node1.Code != 0 && Node2.Code != 0 && ElemCount12 == 1)
                {
                    if (TempInt4 != null)
                    {
                        ForwardVector.Add(Node1.ID.ToString() + "," + TempInt4.ID.ToString(), new Vector() { StartNode = Node1, EndNode = TempInt4 });
                        BackwardVector.Add(TempInt4.ID.ToString() + "," + Node1.ID.ToString(), new Vector() { StartNode = TempInt4, EndNode = Node1 });
                    }
                }

                if (TempInt3 != null && TempInt4 != null)
                {
                    ForwardVector.Add(TempInt3.ID.ToString() + "," + TempInt4.ID.ToString(), new Vector() { StartNode = TempInt3, EndNode = TempInt4 });
                    BackwardVector.Add(TempInt4.ID.ToString() + "," + TempInt3.ID.ToString(), new Vector() { StartNode = TempInt4, EndNode = TempInt3 });
                }

            }
            else if (Node0.Value < ContourValue && Node1.Value < ContourValue && Node2.Value >= ContourValue && Node3.Value >= ContourValue)
            {
                Node TempInt1 = InterpolatedContourNodeList.Where(IntNode => IntNode.ID == Node3.ID * 1000000 + Node0.ID).First();
                if (Node3.Code != 0 && Node0.Code != 0 && ElemCount03 == 1)
                {
                    if (TempInt1 != null)
                    {
                        ForwardVector.Add(Node3.ID.ToString() + "," + TempInt1.ID.ToString(), new Vector() { StartNode = Node3, EndNode = TempInt1 });
                        BackwardVector.Add(TempInt1.ID.ToString() + "," + Node3.ID.ToString(), new Vector() { StartNode = TempInt1, EndNode = Node3 });
                    }


                }
                Node TempInt2 = InterpolatedContourNodeList.Where(IntNode => IntNode.ID == Node2.ID * 1000000 + Node1.ID).First();
                if (Node2.Code != 0 && Node1.Code != 0 && ElemCount12 == 1)
                {
                    if (TempInt2 != null)
                    {
                        ForwardVector.Add(Node2.ID.ToString() + "," + TempInt2.ID.ToString(), new Vector() { StartNode = Node2, EndNode = TempInt2 });
                        BackwardVector.Add(TempInt2.ID.ToString() + "," + Node2.ID.ToString(), new Vector() { StartNode = TempInt2, EndNode = Node2 });
                    }
                }

                if (TempInt1 != null && TempInt2 != null)
                {
                    ForwardVector.Add(TempInt1.ID.ToString() + "," + TempInt2.ID.ToString(), new Vector() { StartNode = TempInt1, EndNode = TempInt2 });
                    BackwardVector.Add(TempInt2.ID.ToString() + "," + TempInt1.ID.ToString(), new Vector() { StartNode = TempInt2, EndNode = TempInt1 });
                }

            }
            else if (Node0.Value >= ContourValue && Node1.Value < ContourValue && Node2.Value < ContourValue && Node3.Value < ContourValue)
            {
                Node TempInt1 = InterpolatedContourNodeList.Where(IntNode => IntNode.ID == Node0.ID * 1000000 + Node1.ID).First();
                if (Node0.Code != 0 && Node1.Code != 0 && ElemCount01 == 1)
                {
                    if (TempInt1 != null)
                    {
                        ForwardVector.Add(Node0.ID.ToString() + "," + TempInt1.ID.ToString(), new Vector() { StartNode = Node0, EndNode = TempInt1 });
                        BackwardVector.Add(TempInt1.ID.ToString() + "," + Node0.ID.ToString(), new Vector() { StartNode = TempInt1, EndNode = Node0 });
                    }


                }
                Node TempInt2 = InterpolatedContourNodeList.Where(IntNode => IntNode.ID == Node0.ID * 1000000 + Node3.ID).First();
                if (Node0.Code != 0 && Node3.Code != 0 && ElemCount03 == 1)
                {
                    if (TempInt2 != null)
                    {
                        ForwardVector.Add(Node0.ID.ToString() + "," + TempInt2.ID.ToString(), new Vector() { StartNode = Node0, EndNode = TempInt2 });
                        BackwardVector.Add(TempInt2.ID.ToString() + "," + Node0.ID.ToString(), new Vector() { StartNode = TempInt2, EndNode = Node0 });
                    }
                }

                if (TempInt1 != null && TempInt2 != null)
                {
                    ForwardVector.Add(TempInt1.ID.ToString() + "," + TempInt2.ID.ToString(), new Vector() { StartNode = TempInt1, EndNode = TempInt2 });
                    BackwardVector.Add(TempInt2.ID.ToString() + "," + TempInt1.ID.ToString(), new Vector() { StartNode = TempInt2, EndNode = TempInt1 });
                }

            }
            else if (Node0.Value < ContourValue && Node1.Value >= ContourValue && Node2.Value < ContourValue && Node3.Value < ContourValue)
            {
                Node TempInt1 = InterpolatedContourNodeList.Where(IntNode => IntNode.ID == Node1.ID * 1000000 + Node0.ID).First();
                if (Node1.Code != 0 && Node0.Code != 0 && ElemCount01 == 1)
                {
                    if (TempInt1 != null)
                    {
                        ForwardVector.Add(Node1.ID.ToString() + "," + TempInt1.ID.ToString(), new Vector() { StartNode = Node1, EndNode = TempInt1 });
                        BackwardVector.Add(TempInt1.ID.ToString() + "," + Node1.ID.ToString(), new Vector() { StartNode = TempInt1, EndNode = Node1 });
                    }


                }
                Node TempInt2 = InterpolatedContourNodeList.Where(IntNode => IntNode.ID == Node1.ID * 1000000 + Node2.ID).First();
                if (Node1.Code != 0 && Node2.Code != 0 && ElemCount12 == 1)
                {
                    if (TempInt2 != null)
                    {
                        ForwardVector.Add(Node1.ID.ToString() + "," + TempInt2.ID.ToString(), new Vector() { StartNode = Node1, EndNode = TempInt2 });
                        BackwardVector.Add(TempInt2.ID.ToString() + "," + Node1.ID.ToString(), new Vector() { StartNode = TempInt2, EndNode = Node1 });
                    }
                }

                if (TempInt1 != null && TempInt2 != null)
                {
                    ForwardVector.Add(TempInt1.ID.ToString() + "," + TempInt2.ID.ToString(), new Vector() { StartNode = TempInt1, EndNode = TempInt2 });
                    BackwardVector.Add(TempInt2.ID.ToString() + "," + TempInt1.ID.ToString(), new Vector() { StartNode = TempInt2, EndNode = TempInt1 });
                }

            }
            else if (Node0.Value < ContourValue && Node1.Value < ContourValue && Node2.Value >= ContourValue && Node3.Value < ContourValue)
            {
                Node TempInt1 = InterpolatedContourNodeList.Where(IntNode => IntNode.ID == Node2.ID * 1000000 + Node1.ID).First();
                if (Node2.Code != 0 && Node1.Code != 0 && ElemCount12 == 1)
                {
                    if (TempInt1 != null)
                    {
                        ForwardVector.Add(Node2.ID.ToString() + "," + TempInt1.ID.ToString(), new Vector() { StartNode = Node2, EndNode = TempInt1 });
                        BackwardVector.Add(TempInt1.ID.ToString() + "," + Node2.ID.ToString(), new Vector() { StartNode = TempInt1, EndNode = Node2 });
                    }


                }
                Node TempInt2 = InterpolatedContourNodeList.Where(IntNode => IntNode.ID == Node2.ID * 1000000 + Node3.ID).First();
                if (Node2.Code != 0 && Node3.Code != 0 && ElemCount23 == 1)
                {
                    if (TempInt2 != null)
                    {
                        ForwardVector.Add(Node2.ID.ToString() + "," + TempInt2.ID.ToString(), new Vector() { StartNode = Node2, EndNode = TempInt2 });
                        BackwardVector.Add(TempInt2.ID.ToString() + "," + Node2.ID.ToString(), new Vector() { StartNode = TempInt2, EndNode = Node2 });
                    }
                }

                if (TempInt1 != null && TempInt2 != null)
                {
                    ForwardVector.Add(TempInt1.ID.ToString() + "," + TempInt2.ID.ToString(), new Vector() { StartNode = TempInt1, EndNode = TempInt2 });
                    BackwardVector.Add(TempInt2.ID.ToString() + "," + TempInt1.ID.ToString(), new Vector() { StartNode = TempInt2, EndNode = TempInt1 });
                }

            }
            else if (Node0.Value < ContourValue && Node1.Value < ContourValue && Node2.Value < ContourValue && Node3.Value >= ContourValue)
            {
                Node TempInt1 = InterpolatedContourNodeList.Where(IntNode => IntNode.ID == Node3.ID * 1000000 + Node0.ID).First();
                if (Node3.Code != 0 && Node0.Code != 0 && ElemCount03 == 1)
                {
                    if (TempInt1 != null)
                    {
                        ForwardVector.Add(Node3.ID.ToString() + "," + TempInt1.ID.ToString(), new Vector() { StartNode = Node3, EndNode = TempInt1 });
                        BackwardVector.Add(TempInt1.ID.ToString() + "," + Node3.ID.ToString(), new Vector() { StartNode = TempInt1, EndNode = Node3 });
                    }


                }
                Node TempInt2 = InterpolatedContourNodeList.Where(IntNode => IntNode.ID == Node3.ID * 1000000 + Node2.ID).First();
                if (Node3.Code != 0 && Node2.Code != 0 && ElemCount23 == 1)
                {
                    if (TempInt2 != null)
                    {
                        ForwardVector.Add(Node3.ID.ToString() + "," + TempInt2.ID.ToString(), new Vector() { StartNode = Node3, EndNode = TempInt2 });
                        BackwardVector.Add(TempInt2.ID.ToString() + "," + Node3.ID.ToString(), new Vector() { StartNode = TempInt2, EndNode = Node3 });
                    }
                }

                if (TempInt1 != null && TempInt2 != null)
                {
                    ForwardVector.Add(TempInt1.ID.ToString() + "," + TempInt2.ID.ToString(), new Vector() { StartNode = TempInt1, EndNode = TempInt2 });
                    BackwardVector.Add(TempInt2.ID.ToString() + "," + TempInt1.ID.ToString(), new Vector() { StartNode = TempInt2, EndNode = TempInt1 });
                }

            }
            else if (Node0.Value < ContourValue && Node1.Value < ContourValue && Node2.Value < ContourValue && Node3.Value < ContourValue)
            {
                // no vector to create
            }
            else
            {
                NotUsed = TaskRunnerServiceRes.AllNodesAreSmallerThanContourValue;
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageList("AllNodesAreSmallerThanContourValue");
                return false;
            }

            return true;
        }
        private string GetFileNameOnlyText(PFSFile pfsFile, string Path, string Keyword)
        {
            string NotUsed = "";
            string FileName = "";

            PFSSection pfsSectionFileName = pfsFile.GetSectionFromHandle(Path);

            if (pfsSectionFileName != null)
            {
                PFSKeyword keyword = null;
                try
                {
                    keyword = pfsSectionFileName.GetKeyword(Keyword);
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
                        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("PFS_Error_", "GetKeyword", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                        return FileName;
                    }
                }
            }

            return FileName;
        }
        public DfsuFile GetHydrodynamicDfsuFile()
        {
            string NotUsed = "";

            DfsuFile dfsuFile = null;
            if (_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID == 0)
            {
                NotUsed = string.Format(TaskRunnerServiceRes._ShouldNotBe0InFunction_, "_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID", "GetHydrodynamicDfsuFile");
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("_ShouldNotBe0InFunction_", "_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID", "GetHydrodynamicDfsuFile");
                return dfsuFile;
            }

            TVFileModel tvFileModelM21_3fm = _TVFileService.GetTVFileModelWithTVItemIDAndTVFileTypeM21FMOrM3FMDB(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID);
            if (!string.IsNullOrWhiteSpace(tvFileModelM21_3fm.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_M21OrM3MDBWith_Equal_, TaskRunnerServiceRes.TVFile, TaskRunnerServiceRes.TVItemID, _TaskRunnerBaseService._BWObj.appTaskModel.TVItemID.ToString());
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_M21OrM3MDBWith_Equal_", TaskRunnerServiceRes.TVFile, TaskRunnerServiceRes.TVItemID, _TaskRunnerBaseService._BWObj.appTaskModel.TVItemID.ToString());
                return dfsuFile;
            }

            PFSFile pfsFile = new PFSFile(tvFileModelM21_3fm.ServerFilePath + tvFileModelM21_3fm.ServerFileName);
            string HydroFileName = GetFileNameOnlyText(pfsFile, "FemEngineHD/HYDRODYNAMIC_MODULE/OUTPUTS/OUTPUT_1", "file_name");
            if (string.IsNullOrWhiteSpace(HydroFileName))
            {
                if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.PFS_Error_, "file_name", "need to specify the error");
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("PFS_Error_", "file_name", "need to specify the error");
                }
                return dfsuFile;
            }

            FileInfo fiHydro = new FileInfo(tvFileModelM21_3fm.ServerFilePath + HydroFileName);
            if (!fiHydro.Exists)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.File_DoesNotExist, fiHydro.FullName);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("File_DoesNotExist", fiHydro.FullName);
                return dfsuFile;
            }

            try
            {
                dfsuFile = DfsuFile.Open(fiHydro.FullName);
            }
            catch (Exception)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotOpenFile_, fiHydro.FullName);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotOpenFile_", fiHydro.FullName);
                return dfsuFile;
            }

            return dfsuFile;
        }
        private string GetScenarioName()
        {
            MikeScenarioModel mikeScenarioModel = _MikeScenarioService.GetMikeScenarioModelWithMikeScenarioTVItemIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID);
            if (string.IsNullOrWhiteSpace(mikeScenarioModel.Error))
            {
                return mikeScenarioModel.MikeScenarioTVText;
            }

            return "";
        }
        private Coord GetSourceCoord(PFSSection pfsSectionSource)
        {
            string NotUsed = "";
            Coord SourceCoord = null;
            PFSKeyword pfsKeywordCoord = null;
            try
            {
                pfsKeywordCoord = pfsSectionSource.GetKeyword("coordinates");
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
                    NotUsed = string.Format(TaskRunnerServiceRes.PFS_Error_, ";", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("PFS_Error_", ";", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                    return SourceCoord;
                }
            }

            return SourceCoord;
        }
        private float? GetSourceFlow(PFSSection pfsSectionSource)
        {
            string NotUsed = "";
            float? SourceFlow = null;
            PFSKeyword pfsKeywordFlow = null;
            try
            {
                pfsKeywordFlow = pfsSectionSource.GetKeyword("constant_value");
            }
            catch (Exception ex)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.PFS_Error_, "GetKeyword", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("PFS_Error_", "GetKeyword", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                return SourceFlow;
            }

            if (pfsKeywordFlow != null)
            {
                try
                {
                    SourceFlow = (float)pfsKeywordFlow.GetParameter(1).ToDouble();
                }
                catch (Exception ex)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.PFS_Error_, ";", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("PFS_Error_", ";", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                    return SourceFlow;
                }
            }

            return SourceFlow;
        }
        private int? GetSourceIncluded(PFSSection pfsSectionSource)
        {
            string NotUsed = "";

            int? SourceIncluded = null;
            PFSKeyword pfsKeywordInculde = null;
            try
            {
                pfsKeywordInculde = pfsSectionSource.GetKeyword("include");
            }
            catch (Exception ex)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.PFS_Error_, "GetKeyword", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("PFS_Error_", "GetKeyword", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                return SourceIncluded;
            }

            if (pfsKeywordInculde != null)
            {
                try
                {
                    SourceIncluded = pfsKeywordInculde.GetParameter(1).ToInt();
                }
                catch (Exception ex)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.PFS_Error_, "GetParameter", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("PFS_Error_", "GetParameter", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                    return SourceIncluded;
                }
            }

            return SourceIncluded;
        }
        private string GetSourceName(PFSSection pfsSectionSource)
        {
            string NotUsed = "";
            string Name = "";
            PFSKeyword pfsKeywordName = null;
            try
            {
                pfsKeywordName = pfsSectionSource.GetKeyword("Name");
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
        private int? GetSourcePollution(PFSSection pfsSectionSourceTransport)
        {
            string NotUsed = "";
            int? SourcePollution = null;
            PFSKeyword pfsKeywordPollution = null;
            try
            {
                pfsKeywordPollution = pfsSectionSourceTransport.GetKeyword("constant_value");
            }
            catch (Exception ex)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.PFS_Error_, "GetKeyword", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("PFS_Error_", "GetKeyword", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                return SourcePollution;
            }

            if (pfsKeywordPollution != null)
            {
                try
                {
                    SourcePollution = (int)pfsKeywordPollution.GetParameter(1).ToInt();
                }
                catch (Exception ex)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.PFS_Error_, "GetParameter", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("PFS_Error_", "GetParameter", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                    return SourcePollution;
                }
            }

            return SourcePollution;
        }
        private int? GetSourcePollutionContinuous(PFSSection pfsSectionSourceTransport)
        {
            string NotUsed = "";
            int? SourcePollutionContinuous = null;
            PFSKeyword pfsKeywordPollutionContinuous = null;
            try
            {
                pfsKeywordPollutionContinuous = pfsSectionSourceTransport.GetKeyword("format");
            }
            catch (Exception ex)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.PFS_Error_, "GetKeyword", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("PFS_Error_", "GetKeyword", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                return SourcePollutionContinuous;
            }

            if (pfsKeywordPollutionContinuous != null)
            {
                try
                {
                    SourcePollutionContinuous = (int)pfsKeywordPollutionContinuous.GetParameter(1).ToInt();
                }
                catch (Exception ex)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.PFS_Error_, "GetParameter", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("PFS_Error_", "GetParameter", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                    return SourcePollutionContinuous;
                }
            }

            return SourcePollutionContinuous;
        }
        private float? GetSourceSalinity(PFSSection pfsSectionSourceSalinity)
        {
            string NotUsed = "";
            float? SourceSalinity = null;
            PFSKeyword pfsKeywordSalinity = null;
            try
            {
                pfsKeywordSalinity = pfsSectionSourceSalinity.GetKeyword("constant_value");
            }
            catch (Exception ex)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.PFS_Error_, "GetKeyword", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("PFS_Error_", "GetKeyword", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                return SourceSalinity;
            }

            if (pfsKeywordSalinity != null)
            {
                try
                {
                    SourceSalinity = (float)pfsKeywordSalinity.GetParameter(1).ToDouble();
                }
                catch (Exception ex)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.PFS_Error_, "GetParameter", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("PFS_Error_", "GetParameter", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                    return SourceSalinity;
                }
            }

            return SourceSalinity;
        }
        private float? GetSourceTemperature(PFSSection pfsSectionSourceTemperature)
        {
            string NotUsed = "";
            float? SourceTemperature = null;
            PFSKeyword pfsKeywordTemperature = null;
            try
            {
                pfsKeywordTemperature = pfsSectionSourceTemperature.GetKeyword("constant_value");
            }
            catch (Exception ex)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.PFS_Error_, "GetKeyword", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("PFS_Error_", "GetKeyword", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                return SourceTemperature;
            }

            if (pfsKeywordTemperature != null)
            {
                try
                {
                    SourceTemperature = (float)pfsKeywordTemperature.GetParameter(1).ToDouble();
                }
                catch (Exception ex)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.PFS_Error_, "GetParameter", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("PFS_Error_", "GetParameter", ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
                    return SourceTemperature;
                }
            }

            return SourceTemperature;
        }
        private DfsuFile GetTransportDfsuFile()
        {
            string NotUsed = "";

            DfsuFile dfsuFile = null;
            if (_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID == 0)
            {
                NotUsed = string.Format(TaskRunnerServiceRes._ShouldNotBe0InFunction_, "_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID", "GetHydrodynamicDfsuFile");
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("_ShouldNotBe0InFunction_", "_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID", "GetHydrodynamicDfsuFile");
                return dfsuFile;
            }

            TVFileModel tvFileModelM21_3fm = _TVFileService.GetTVFileModelWithTVItemIDAndTVFileTypeM21FMOrM3FMDB(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID);
            if (!string.IsNullOrWhiteSpace(tvFileModelM21_3fm.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_M21OrM3MDBWith_Equal_, TaskRunnerServiceRes.TVFile, TaskRunnerServiceRes.TVItemID, _TaskRunnerBaseService._BWObj.appTaskModel.TVItemID);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_M21OrM3MDBWith_Equal_", TaskRunnerServiceRes.TVFile, TaskRunnerServiceRes.TVItemID, _TaskRunnerBaseService._BWObj.appTaskModel.TVItemID.ToString());
                return dfsuFile;
            }

            PFSFile pfsFile = new PFSFile(tvFileModelM21_3fm.ServerFilePath + tvFileModelM21_3fm.ServerFileName);
            string TransFileName = GetFileNameOnlyText(pfsFile, "FemEngineHD/TRANSPORT_MODULE/OUTPUTS/OUTPUT_1", "file_name");
            if (string.IsNullOrWhiteSpace(TransFileName))
            {
                if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.PFS_Error_, "file_name", "need to fill the error message");
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("PFS_Error_", "file_name", "need to fill the error message");
                }
                return dfsuFile;
            }

            FileInfo fiTrans = new FileInfo(tvFileModelM21_3fm.ServerFilePath + TransFileName);
            if (!fiTrans.Exists)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.File_DoesNotExist, fiTrans.FullName);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("File_DoesNotExist", fiTrans.FullName);
                return dfsuFile;
            }

            try
            {
                dfsuFile = DfsuFile.Open(fiTrans.FullName);
            }
            catch (Exception)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotOpenFile_, fiTrans.FullName);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotOpenFile_", fiTrans.FullName);
                return dfsuFile;
            }

            return dfsuFile;
        }
        private void InsertNewNodeInInterpolatedContourNodeList(DfsuFile dfsuFile, Node NodeLarge, Node NodeSmall, float ContourValue)
        {
            PolyPoint point = new PolyPoint();
            point.XCoord = NodeSmall.X + (NodeLarge.X - NodeSmall.X) * (ContourValue - NodeSmall.Value) / (NodeLarge.Value - NodeSmall.Value);
            point.YCoord = NodeSmall.Y + (NodeLarge.Y - NodeSmall.Y) * (ContourValue - NodeSmall.Value) / (NodeLarge.Value - NodeSmall.Value);
            point.Z = dfsuFile.Z[NodeSmall.ID - 1] + (dfsuFile.Z[NodeLarge.ID - 1] - dfsuFile.Z[NodeSmall.ID - 1]) * (ContourValue - NodeSmall.Value) / (NodeLarge.Value - NodeSmall.Value);

            Node NewNode = new Node();
            NewNode.ID = 1000000 * NodeLarge.ID + NodeSmall.ID;
            NewNode.X = point.XCoord;
            NewNode.Y = point.YCoord;
            NewNode.Z = point.Z;
            NewNode.Value = ContourValue;

            if (InterpolatedContourNodeList.Where(nn => nn.ID == NewNode.ID).Count() == 0)
            {
                InterpolatedContourNodeList.Add(NewNode);
            }
        }
        public string UpdateAppTaskPercentCompleted(int AppTaskID, int PercentCompleted)
        {
            if (_TaskRunnerBaseService._User != null)
            {
                using (AppTaskService appTaskService = new AppTaskService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User))
                {
                    AppTaskModel appTaskModel = appTaskService.GetAppTaskModelWithAppTaskIDDB(AppTaskID);
                    if (!string.IsNullOrWhiteSpace(appTaskModel.Error))
                        return appTaskModel.Error;

                    appTaskModel.PercentCompleted = PercentCompleted;
                    AppTaskModel appTaskModelRet = appTaskService.PostUpdateAppTask(appTaskModel);
                    if (!string.IsNullOrWhiteSpace(appTaskModelRet.Error))
                        return appTaskModel.Error;
                }
            }

            return "";
        }
        private bool WriteKMLBottom(StringBuilder sbHTML)
        {
            sbHTML.AppendLine(@"</Document>");
            sbHTML.AppendLine(@"</kml>");

            return true;
        }
        private bool WriteKMLBoundaryConditionNode(StringBuilder sbHTML)
        {
            string[] Colors = { "ylw", "grn", "blue", "ltblu", "pink", "red" };

            foreach (string color in Colors)
            {
                sbHTML.AppendLine(string.Format(@"	<Style id=""sn_{0}-pushpin"">", color));
                sbHTML.AppendLine(@"		<IconStyle>");
                sbHTML.AppendLine(@"			<scale>1.1</scale>");
                sbHTML.AppendLine(@"			<Icon>");
                sbHTML.AppendLine(string.Format(@"				<href>http://maps.google.com/mapfiles/kml/pushpin/{0}-pushpin.png</href>", color));
                sbHTML.AppendLine(@"			</Icon>");
                sbHTML.AppendLine(@"			<hotSpot x=""20"" y=""2"" xunits=""pixels"" yunits=""pixels""/>");
                sbHTML.AppendLine(@"		</IconStyle>");
                sbHTML.AppendLine(@"		<ListStyle>");
                sbHTML.AppendLine(@"		</ListStyle>");
                sbHTML.AppendLine(@"	</Style>");
                sbHTML.AppendLine(string.Format(@"	<StyleMap id=""msn_{0}-pushpin"">", color));
                sbHTML.AppendLine(@"		<Pair>");
                sbHTML.AppendLine(@"			<key>normal</key>");
                sbHTML.AppendLine(string.Format(@"			<styleUrl>#sn_{0}-pushpin</styleUrl>", color));
                sbHTML.AppendLine(@"		</Pair>");
                sbHTML.AppendLine(@"		<Pair>");
                sbHTML.AppendLine(@"			<key>highlight</key>");
                sbHTML.AppendLine(string.Format(@"			<styleUrl>#sh_{0}-pushpin</styleUrl>", color));
                sbHTML.AppendLine(@"		</Pair>");
                sbHTML.AppendLine(@"	</StyleMap>");
                sbHTML.AppendLine(string.Format(@"	<Style id=""sh_{0}-pushpin"">", color));
                sbHTML.AppendLine(@"		<IconStyle>");
                sbHTML.AppendLine(@"			<scale>1.3</scale>");
                sbHTML.AppendLine(@"			<Icon>");
                sbHTML.AppendLine(string.Format(@"				<href>http://maps.google.com/mapfiles/kml/pushpin/{0}-pushpin.png</href>", color));
                sbHTML.AppendLine(@"			</Icon>");
                sbHTML.AppendLine(@"			<hotSpot x=""20"" y=""2"" xunits=""pixels"" yunits=""pixels""/>");
                sbHTML.AppendLine(@"		</IconStyle>");
                sbHTML.AppendLine(@"		<ListStyle>");
                sbHTML.AppendLine(@"		</ListStyle>");
                sbHTML.AppendLine(@"	</Style>");
            }

            //UpdateTask(AppTaskID, "30 %");

            TVItemModel tvItemModelMikeScenario = _TVItemService.GetTVItemModelWithTVItemIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID);
            if (!string.IsNullOrWhiteSpace(tvItemModelMikeScenario.Error))
            {
                sbHTML.AppendLine(@"<Folder>");
                sbHTML.AppendLine(@"    <name>" + TaskRunnerServiceRes.Error + "</name>");
                sbHTML.AppendLine(@"    <description><![CDATA[");
                sbHTML.AppendLine(@"    <h4>" + tvItemModelMikeScenario.Error + "</h4");
                sbHTML.AppendLine(@"    ]]></description>");
                sbHTML.AppendLine(@"</Folder>");
                return true;
            }

            List<MikeBoundaryConditionModel> mbcModelList = _MikeBoundaryConditionService.GetMikeBoundaryConditionModelListWithMikeScenarioTVItemIDAndTVTypeDB(tvItemModelMikeScenario.TVItemID, TVTypeEnum.MikeBoundaryConditionMesh);

            int countColor = 0;
            foreach (MikeBoundaryConditionModel mbcm in mbcModelList)
            {
                sbHTML.AppendLine(@"<Folder>");
                sbHTML.AppendLine(@"<name>" + mbcm.MikeBoundaryConditionName + " (" + mbcm.MikeBoundaryConditionCode + ") length [" + mbcm.MikeBoundaryConditionLength_m.ToString("F0") + "] </name>");
                sbHTML.AppendLine(@"<visibility>1</visibility>");
                sbHTML.AppendLine(@"<description><![CDATA[");
                sbHTML.AppendLine(@"<p>(" + mbcm.MikeBoundaryConditionFormat + ") " + mbcm.MikeBoundaryConditionLevelOrVelocity + " " + mbcm.WebTideDataSet.ToString() + " " + mbcm.NumberOfWebTideNodes + " " + TaskRunnerServiceRes.Nodes + "</p>");
                sbHTML.AppendLine(@"]]></description>");

                // drawing Boundary Nodes
                List<MapInfoPointModel> mapInfoPointModelList = _MapInfoService._MapInfoPointService.GetMapInfoPointModelListWithTVItemIDAndTVTypeAndMapInfoDrawTypeDB(mbcm.MikeBoundaryConditionTVItemID, TVTypeEnum.MikeBoundaryConditionMesh, MapInfoDrawTypeEnum.Polyline);

                sbHTML.AppendLine(@"    <Folder>");
                sbHTML.AppendLine(@"    <name>" + TaskRunnerServiceRes.ElementNodes + @"</name>");
                sbHTML.AppendLine(@"    <open>1</open>");
                foreach (MapInfoPointModel mapInfoPointModel in mapInfoPointModelList)
                {
                    sbHTML.AppendLine(@"    <Placemark>");
                    sbHTML.AppendLine(@"    <name>Node " + mapInfoPointModel.Ordinal + "</name>");
                    sbHTML.AppendLine(string.Format(@"    <styleUrl>#msn_{0}-pushpin</styleUrl>", Colors[countColor]));
                    sbHTML.AppendLine(@"    <Point>");
                    sbHTML.AppendLine(@"    <coordinates>" + mapInfoPointModel.Lng.ToString().Replace(",", ".") + @"," + mapInfoPointModel.Lat.ToString().Replace(",", ".") + @",0</coordinates>");
                    sbHTML.AppendLine(@"    </Point>");
                    sbHTML.AppendLine(@"    </Placemark>");
                }
                sbHTML.AppendLine(@"    </Folder>");


                countColor += 1;
                if (countColor > 5) countColor = 0;
                MikeBoundaryConditionModel mbcModel2 = _MikeBoundaryConditionService.GetMikeBoundaryConditionModelListWithMikeScenarioTVItemIDAndTVTypeDB(tvItemModelMikeScenario.TVItemID, TVTypeEnum.MikeBoundaryConditionWebTide).Where(c => c.MikeBoundaryConditionName == mbcm.MikeBoundaryConditionName).FirstOrDefault();
                List<MapInfoPointModel> mapInfoPointModelList2 = _MapInfoService._MapInfoPointService.GetMapInfoPointModelListWithTVItemIDAndTVTypeAndMapInfoDrawTypeDB(mbcModel2.MikeBoundaryConditionTVItemID, TVTypeEnum.MikeBoundaryConditionWebTide, MapInfoDrawTypeEnum.Polyline);

                sbHTML.AppendLine(@"    <Folder>");
                sbHTML.AppendLine(@"    <name>" + TaskRunnerServiceRes.WebTideNodes + @"</name>");
                sbHTML.AppendLine(@"    <open>1</open>");
                foreach (MapInfoPointModel mapInfoPointModel in mapInfoPointModelList2)
                {
                    sbHTML.AppendLine(@"    <Placemark>");
                    sbHTML.AppendLine(@"    <name>Node " + mapInfoPointModel.Ordinal + "</name>");
                    sbHTML.AppendLine(string.Format(@"    <styleUrl>#msn_{0}-pushpin</styleUrl>", Colors[countColor]));
                    sbHTML.AppendLine(@"    <Point>");
                    sbHTML.AppendLine(@"    <coordinates>" + mapInfoPointModel.Lng.ToString().Replace(",", ".") + @"," + mapInfoPointModel.Lat.ToString().Replace(",", ".") + @",0</coordinates>");
                    sbHTML.AppendLine(@"    </Point>");
                    sbHTML.AppendLine(@"    </Placemark>");
                }
                sbHTML.AppendLine(@"    </Folder>");

                sbHTML.AppendLine(@"</Folder>");

                countColor += 1;
                if (countColor > 5) countColor = 0;
            }

            return true;
        }
        private bool WriteKMLFecalColiformContourLine(DfsuFile dfsuFile, List<ElementLayer> elementLayerList, List<NodeLayer> topNodeLayerList, List<NodeLayer> bottomNodeLayerList, List<float> ContourValueList)
        {
            string NotUsed = "";
            int PercentCompleted = 3;
            int ItemNumber = 0;
            double RefreshEveryXSeconds = double.Parse("5");
            DateTime RefreshDateTime = DateTime.Now.AddSeconds(RefreshEveryXSeconds);

            // getting the ItemNumber
            foreach (IDfsSimpleDynamicItemInfo dfsDyInfo in dfsuFile.ItemInfo)
            {
                if (dfsDyInfo.Quantity.Item == eumItem.eumIConcentration
                     || dfsDyInfo.Quantity.Item == eumItem.eumIConcentration1
                     || dfsDyInfo.Quantity.Item == eumItem.eumIConcentration_1
                     || dfsDyInfo.Quantity.Item == eumItem.eumIConcentration_2
                     || dfsDyInfo.Quantity.Item == eumItem.eumIConcentration_3
                     || dfsDyInfo.Quantity.Item == eumItem.eumIConcentration_4)
                {
                    ItemNumber = dfsDyInfo.ItemNumber;
                    break;
                }
            }

            if (ItemNumber == 0)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind__, TaskRunnerServiceRes.ParameterType, "eumIConcentration");
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotFind__", TaskRunnerServiceRes.ParameterType, "eumIConcentration");
                return false;
            }

            //            int pcount = 0;
            sb.AppendLine(@"<Folder>");
            sb.AppendLine(@"  <name>" + TaskRunnerServiceRes.PollutionAnimation + "</name>");
            sb.AppendLine(@"  <visibility>0</visibility>");

            int CountAt = 0;
            int CountLayer = (dfsuFile.NumberOfSigmaLayers == 0 ? 1 : dfsuFile.NumberOfSigmaLayers);
            int CurrentContourValue = 0;
            int CurrentTimeSteps = 0;

            int TotalCount = CountLayer * ContourValueList.Count * dfsuFile.NumberOfTimeSteps;

            for (int Layer = 1; Layer <= CountLayer; Layer++)
            {
                //CurrentLayer += 1;
                CurrentContourValue = 1;
                CurrentTimeSteps = 1;

                #region Top of Layer
                if (Layer == 1)
                {
                    sb.AppendLine(string.Format(@"<Folder><name>" + TaskRunnerServiceRes.TopOfLayer + @" [{0}] </name>", Layer));
                }
                else
                {
                    sb.AppendLine(string.Format(@"<Folder><name>" + TaskRunnerServiceRes.TopOfLayer + @" [{0}] " + TaskRunnerServiceRes.BottomOfLayer + @" [{1}] </name>", Layer, Layer - 1));
                }
                sb.AppendLine(@"<visibility>0</visibility>");
                int CountContourValue = 1;
                foreach (float ContourValue in ContourValueList)
                {
                    sb.AppendLine(string.Format(@"  <Folder><name>" + TaskRunnerServiceRes.ContourValue + @" [{0}]</name>", ContourValue));
                    sb.AppendLine(@"  <visibility>0</visibility>");

                    int vcount = 0;
                    CurrentContourValue += 1;
                    //for (int timeStep = 30; timeStep < 35 /*dfsuFile.NumberOfTimeSteps */; timeStep++)
                    for (int timeStep = 0; timeStep < dfsuFile.NumberOfTimeSteps; timeStep++)
                    {
                        CountAt += 1;
                        int PercentCompletedTemp = (int)(((CountAt - 1) * (float)100.0f) / (float)TotalCount);
                        if (PercentCompleted != PercentCompletedTemp)
                        {
                            string retStr = UpdateAppTaskPercentCompleted(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, PercentCompleted);
                        }
                        PercentCompleted = PercentCompletedTemp;

                        float[] ValueList = (float[])dfsuFile.ReadItemTimeStep(ItemNumber, timeStep).Data;

                        List<ContourPolygon> ContourPolygonList = new List<ContourPolygon>();

                        for (int i = 0; i < elementLayerList.Count; i++)
                        {
                            elementLayerList[i].Element.Value = ValueList[i];
                        }

                        List<ElementLayer> elementLayerListAtLayer = elementLayerList.Where(c => c.Layer == Layer).ToList();
                        List<NodeLayer> topNodeLayerListAtLayer = topNodeLayerList.Where(c => c.Layer == Layer).ToList();

                        foreach (NodeLayer nl in topNodeLayerListAtLayer)
                        {
                            float Total = 0;
                            foreach (Element element in nl.Node.ElementList)
                            {
                                Total += element.Value;
                            }
                            nl.Node.Value = Total / nl.Node.ElementList.Count;
                        }

                        List<Node> AllNodeList = new List<Node>();

                        List<NodeLayer> AboveNodeLayerList = (from n in topNodeLayerListAtLayer
                                                              where (n.Node.Value >= ContourValue)
                                                              && n.Layer == Layer
                                                              select n).ToList<NodeLayer>();

                        foreach (NodeLayer snl in AboveNodeLayerList)
                        {
                            List<NodeLayer> EndNodeLayerList = null;

                            List<NodeLayer> NodeLayerConnectedList = (from nll in topNodeLayerListAtLayer
                                                                      from n in snl.Node.ConnectNodeList
                                                                      where (n.ID == nll.Node.ID)
                                                                      select nll).ToList<NodeLayer>();

                            EndNodeLayerList = (from nll in NodeLayerConnectedList
                                                where (nll.Node.ID != snl.Node.ID)
                                                && (nll.Node.Value < ContourValue)
                                                && nll.Layer == Layer
                                                select nll).ToList<NodeLayer>();

                            foreach (NodeLayer en in EndNodeLayerList)
                            {
                                AllNodeList.Add(en.Node);
                            }

                            if (snl.Node.Code != 0)
                            {
                                AllNodeList.Add(snl.Node);
                            }

                        }

                        //if (AllNodeList.Count == 0)
                        //{
                        //    //vcount += 1;
                        //    continue;
                        //}

                        List<Element> TempUniqueElementList = new List<Element>();
                        List<Element> UniqueElementList = new List<Element>();
                        foreach (ElementLayer el in elementLayerListAtLayer)
                        {
                            if (dfsuFile.DfsuFileType == DfsuFileType.Dfsu3DSigma || dfsuFile.DfsuFileType == DfsuFileType.Dfsu3DSigmaZ)
                            {
                                if (el.Element.Type == 32)
                                {
                                    bool NodeBigger = false;
                                    for (int i = 3; i < 6; i++)
                                    {
                                        if (el.Element.NodeList[i].Value >= ContourValue)
                                        {
                                            NodeBigger = true;
                                            break;
                                        }
                                    }
                                    if (NodeBigger)
                                    {
                                        int countTrue = 0;
                                        for (int i = 3; i < 6; i++)
                                        {
                                            if (el.Element.NodeList[i].Value >= ContourValue && el.Element.NodeList[i].Code == 0)
                                            {
                                                countTrue += 1;
                                            }
                                        }
                                        if (countTrue != el.Element.NodeList.Count)
                                        {
                                            TempUniqueElementList.Add(el.Element);
                                        }
                                    }
                                }
                                else if (el.Element.Type == 33)
                                {
                                    bool NodeBigger = false;
                                    for (int i = 4; i < 8; i++)
                                    {
                                        if (el.Element.NodeList[i].Value >= ContourValue)
                                        {
                                            NodeBigger = true;
                                            break;
                                        }
                                    }
                                    if (NodeBigger)
                                    {
                                        int countTrue = 0;
                                        for (int i = 4; i < 8; i++)
                                        {
                                            if (el.Element.NodeList[i].Value >= ContourValue && el.Element.NodeList[i].Code == 0)
                                            {
                                                countTrue += 1;
                                            }
                                        }
                                        if (countTrue != el.Element.NodeList.Count)
                                        {
                                            TempUniqueElementList.Add(el.Element);
                                        }
                                    }
                                }
                                else
                                {
                                    NotUsed = string.Format(TaskRunnerServiceRes.ElementType_IsNotSupported, el.Element.Type.ToString());
                                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("ElementType_IsNotSupported", el.Element.Type.ToString());
                                    return false;
                                }
                            }
                            else if (dfsuFile.DfsuFileType == DfsuFileType.Dfsu2D)
                            {
                                bool NodeBigger = false;
                                for (int i = 0; i < el.Element.NodeList.Count; i++)
                                {
                                    if (el.Element.NodeList[i].Value >= ContourValue)
                                    {
                                        NodeBigger = true;
                                        break;
                                    }
                                }
                                if (NodeBigger)
                                {
                                    int countTrue = 0;
                                    for (int i = 0; i < el.Element.NodeList.Count; i++)
                                    {
                                        if (el.Element.NodeList[i].Value >= ContourValue && el.Element.NodeList[i].Code == 0)
                                        {
                                            countTrue += 1;
                                        }
                                    }
                                    if (countTrue != el.Element.NodeList.Count)
                                    {
                                        TempUniqueElementList.Add(el.Element);
                                    }
                                }
                            }
                        }

                        UniqueElementList = (from el in TempUniqueElementList select el).Distinct().ToList<Element>();

                        // filling InterpolatedContourNodeList
                        InterpolatedContourNodeList = new List<Node>();

                        foreach (Element el in UniqueElementList)
                        {
                            if (dfsuFile.DfsuFileType == DfsuFileType.Dfsu3DSigma || dfsuFile.DfsuFileType == DfsuFileType.Dfsu3DSigmaZ)
                            {
                                if (el.Type == 32)
                                {
                                    if (el.NodeList[3].Value >= ContourValue && el.NodeList[4].Value < ContourValue)
                                    {
                                        InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[3], el.NodeList[4], ContourValue);
                                    }
                                    if (el.NodeList[3].Value >= ContourValue && el.NodeList[5].Value < ContourValue)
                                    {
                                        InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[3], el.NodeList[5], ContourValue);
                                    }
                                    if (el.NodeList[4].Value >= ContourValue && el.NodeList[3].Value < ContourValue)
                                    {
                                        InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[4], el.NodeList[3], ContourValue);
                                    }
                                    if (el.NodeList[4].Value >= ContourValue && el.NodeList[5].Value < ContourValue)
                                    {
                                        InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[4], el.NodeList[5], ContourValue);
                                    }
                                    if (el.NodeList[5].Value >= ContourValue && el.NodeList[4].Value < ContourValue)
                                    {
                                        InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[5], el.NodeList[4], ContourValue);
                                    }
                                    if (el.NodeList[5].Value >= ContourValue && el.NodeList[3].Value < ContourValue)
                                    {
                                        InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[5], el.NodeList[3], ContourValue);
                                    }
                                }
                                else if (el.Type == 33)
                                {
                                    if (el.NodeList[4].Value >= ContourValue && el.NodeList[5].Value < ContourValue)
                                    {
                                        InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[4], el.NodeList[5], ContourValue);
                                    }
                                    if (el.NodeList[4].Value >= ContourValue && el.NodeList[7].Value < ContourValue)
                                    {
                                        InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[4], el.NodeList[7], ContourValue);
                                    }
                                    if (el.NodeList[5].Value >= ContourValue && el.NodeList[4].Value < ContourValue)
                                    {
                                        InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[5], el.NodeList[4], ContourValue);
                                    }
                                    if (el.NodeList[5].Value >= ContourValue && el.NodeList[6].Value < ContourValue)
                                    {
                                        InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[5], el.NodeList[6], ContourValue);
                                    }
                                    if (el.NodeList[6].Value >= ContourValue && el.NodeList[5].Value < ContourValue)
                                    {
                                        InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[6], el.NodeList[5], ContourValue);
                                    }
                                    if (el.NodeList[6].Value >= ContourValue && el.NodeList[7].Value < ContourValue)
                                    {
                                        InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[6], el.NodeList[7], ContourValue);
                                    }
                                    if (el.NodeList[7].Value >= ContourValue && el.NodeList[4].Value < ContourValue)
                                    {
                                        InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[7], el.NodeList[4], ContourValue);
                                    }
                                    if (el.NodeList[7].Value >= ContourValue && el.NodeList[6].Value < ContourValue)
                                    {
                                        InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[7], el.NodeList[6], ContourValue);
                                    }
                                }
                                else
                                {
                                    NotUsed = string.Format(TaskRunnerServiceRes.ElementType_IsNotSupported, el.Type.ToString());
                                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("ElementType_IsNotSupported", el.Type.ToString());
                                    return false;
                                }
                            }
                            else if (dfsuFile.DfsuFileType == DfsuFileType.Dfsu2D)
                            {
                                if (el.Type == 21)
                                {
                                    if (el.NodeList[0].Value >= ContourValue && el.NodeList[1].Value < ContourValue)
                                    {
                                        InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[0], el.NodeList[1], ContourValue);
                                    }
                                    if (el.NodeList[0].Value >= ContourValue && el.NodeList[2].Value < ContourValue)
                                    {
                                        InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[0], el.NodeList[2], ContourValue);
                                    }
                                    if (el.NodeList[1].Value >= ContourValue && el.NodeList[0].Value < ContourValue)
                                    {
                                        InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[1], el.NodeList[0], ContourValue);
                                    }
                                    if (el.NodeList[1].Value >= ContourValue && el.NodeList[2].Value < ContourValue)
                                    {
                                        InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[1], el.NodeList[2], ContourValue);
                                    }
                                    if (el.NodeList[2].Value >= ContourValue && el.NodeList[1].Value < ContourValue)
                                    {
                                        InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[2], el.NodeList[1], ContourValue);
                                    }
                                    if (el.NodeList[2].Value >= ContourValue && el.NodeList[0].Value < ContourValue)
                                    {
                                        InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[2], el.NodeList[0], ContourValue);
                                    }
                                }
                                else if (el.Type == 24)
                                {
                                }
                                else if (el.Type == 25)
                                {
                                    if (el.NodeList[0].Value >= ContourValue && el.NodeList[1].Value < ContourValue)
                                    {
                                        InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[0], el.NodeList[1], ContourValue);
                                    }
                                    if (el.NodeList[0].Value >= ContourValue && el.NodeList[3].Value < ContourValue)
                                    {
                                        InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[0], el.NodeList[3], ContourValue);
                                    }
                                    if (el.NodeList[1].Value >= ContourValue && el.NodeList[0].Value < ContourValue)
                                    {
                                        InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[1], el.NodeList[0], ContourValue);
                                    }
                                    if (el.NodeList[1].Value >= ContourValue && el.NodeList[2].Value < ContourValue)
                                    {
                                        InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[1], el.NodeList[2], ContourValue);
                                    }
                                    if (el.NodeList[2].Value >= ContourValue && el.NodeList[1].Value < ContourValue)
                                    {
                                        InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[2], el.NodeList[1], ContourValue);
                                    }
                                    if (el.NodeList[2].Value >= ContourValue && el.NodeList[3].Value < ContourValue)
                                    {
                                        InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[2], el.NodeList[3], ContourValue);
                                    }
                                    if (el.NodeList[3].Value >= ContourValue && el.NodeList[0].Value < ContourValue)
                                    {
                                        InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[3], el.NodeList[0], ContourValue);
                                    }
                                    if (el.NodeList[3].Value >= ContourValue && el.NodeList[2].Value < ContourValue)
                                    {
                                        InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[3], el.NodeList[2], ContourValue);
                                    }
                                }
                                else
                                {
                                    NotUsed = string.Format(TaskRunnerServiceRes.ElementType_IsNotSupported, el.Type.ToString());
                                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("ElementType_IsNotSupported", el.Type.ToString());
                                    return false;
                                }
                            }
                        }

                        List<Node> UniqueNodeList = (from n in AllNodeList orderby n.ID select n).Distinct().ToList<Node>();

                        // ------------------------- new code --------------------------
                        //                     

                        ForwardVector = new Dictionary<string, Vector>();
                        BackwardVector = new Dictionary<string, Vector>();

                        foreach (Element el in UniqueElementList)
                        {
                            if (el.Type == 21)
                            {
                                FillVectors21_32(el, UniqueElementList, ContourValue, false, true);
                            }
                            else if (el.Type == 24)
                            {
                                NotUsed = TaskRunnerServiceRes.AllNodesAreSmallerThanContourValue;
                                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageList("AllNodesAreSmallerThanContourValue");
                                return false;
                            }
                            else if (el.Type == 25)
                            {
                                FillVectors25_33(el, UniqueElementList, ContourValue, false, true);
                            }
                            else if (el.Type == 32)
                            {
                                FillVectors21_32(el, UniqueElementList, ContourValue, true, true);
                            }
                            else if (el.Type == 33)
                            {
                                FillVectors25_33(el, UniqueElementList, ContourValue, true, true);
                            }
                            else
                            {
                                NotUsed = string.Format(TaskRunnerServiceRes.ElementType_IsNotSupported, el.Type.ToString());
                                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("ElementType_IsNotSupported", el.Type.ToString());
                                return false;
                            }

                        }

                        //-------------- new code ------------------------



                        bool MoreContourLine = true;
                        while (MoreContourLine && ForwardVector.Count > 0)
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

                            if (_MapInfoService.CalculateAreaOfPolygon(FinalContourNodeList) < 0)
                            {
                                FinalContourNodeList.Reverse();
                            }

                            FinalContourNodeList.Add(FinalContourNodeList[0]);
                            ContourPolygon contourPolygon = new ContourPolygon() { };
                            contourPolygon.ContourNodeList = FinalContourNodeList;
                            contourPolygon.ContourValue = ContourValue;
                            contourPolygon.Layer = Layer;
                            ContourPolygonList.Add(contourPolygon);

                            if (ForwardVector.Count == 0)
                            {
                                MoreContourLine = false;
                            }

                        }
                        DrawKMLContourPolygon(ContourPolygonList, dfsuFile, vcount);

                        vcount += 1;
                    }
                    sb.AppendLine(@"  </Folder>");
                    CountContourValue += 1;
                }
                sb.AppendLine(@"  </Folder>");
                #endregion Top of Layer

                #region Bottom of Layer
                //// doing the bottom layer if the current layer is == NumberOfSigmaLayers
                //if (Layer == dfsuFile.NumberOfSigmaLayers)
                //{
                //    sbPlacemarkFeacalColiformContour.AppendLine(string.Format(@"<Folder><name>Bottom of Layer [{0}]</name>", Layer));
                //    sbPlacemarkFeacalColiformContour.AppendLine(@"<visibility>0</visibility>");
                //    CountContourValue = 1;
                //    foreach (float ContourValue in ContourValueList)
                //    {
                //        sbPlacemarkFeacalColiformContour.AppendLine(string.Format(@"<Folder><name>Contour Value [{0}]</name>", ContourValue));
                //        sbPlacemarkFeacalColiformContour.AppendLine(@"<visibility>0</visibility>");

                //        int vcount = 0;
                //        //for (int timeStep = 30; timeStep < 35 /*dfsuFile.NumberOfTimeSteps */; timeStep++)
                //        for (int timeStep = 0; timeStep < dfsuFile.NumberOfTimeSteps; timeStep++)
                //        {
                //            CountRefresh += 1;
                //            CountAt += 1;
                //            if (CountRefresh > UpdateAfter)
                //            {
                //                string AppTaskStatus = "";
                //                if (SigmaLayerValueList.Contains(dfsuFile.NumberOfSigmaLayers))
                //                {
                //                    AppTaskStatus = ((int)((CountAt * 100) / (dfsuFile.NumberOfTimeSteps * (SigmaLayerValueList.Count + 1) * ContourValueList.Count))).ToString() + " %";
                //                }
                //                else
                //                {
                //                    AppTaskStatus = ((int)((CountAt * 100) / (dfsuFile.NumberOfTimeSteps * SigmaLayerValueList.Count * ContourValueList.Count))).ToString() + " %";
                //                }
                //                UpdateTask(AppTaskID, AppTaskStatus);
                //                CountRefresh = 0;
                //            }

                //            float[] ValueList = (float[])dfsuFile.ReadItemTimeStep(ItemNumber, timeStep).Data;

                //            List<ContourPolygon> ContourPolygonList = new List<ContourPolygon>();

                //            for (int i = 0; i < ElementLayerList.Count; i++)
                //            {
                //                ElementLayerList[i].Element.Value = ValueList[i];
                //            }

                //            foreach (NodeLayer nl in BottomNodeLayerList)
                //            {
                //                float Total = 0;
                //                foreach (Element element in nl.Node.ElementList)
                //                {
                //                    Total += element.Value;
                //                }
                //                nl.Node.Value = Total / nl.Node.ElementList.Count;
                //            }


                //            List<Node> AllNodeList = new List<Node>();

                //            List<NodeLayer> AboveNodeLayerList = new List<NodeLayer>();

                //            AboveNodeLayerList = (from n in BottomNodeLayerList
                //                                  where (n.Node.Value >= ContourValue)
                //                                  && n.Layer == Layer
                //                                  select n).ToList<NodeLayer>();

                //            foreach (NodeLayer snl in AboveNodeLayerList)
                //            {
                //                List<NodeLayer> EndNodeLayerList = null;

                //                List<NodeLayer> NodeLayerConnectedList = (from nll in BottomNodeLayerList
                //                                                          from n in snl.Node.ConnectNodeList
                //                                                          where (n.ID == nll.Node.ID)
                //                                                          select nll).ToList<NodeLayer>();

                //                EndNodeLayerList = (from nll in NodeLayerConnectedList
                //                                    where (nll.Node.ID != snl.Node.ID)
                //                                    && (nll.Node.Value < ContourValue)
                //                                    && nll.Layer == Layer
                //                                    select nll).ToList<NodeLayer>();

                //                foreach (NodeLayer en in EndNodeLayerList)
                //                {
                //                    AllNodeList.Add(en.Node);
                //                }

                //                if (snl.Node.Code != 0)
                //                {
                //                    AllNodeList.Add(snl.Node);
                //                }

                //            }

                //            if (AllNodeList.Count == 0)
                //            {
                //                //vcount += 1;
                //                continue;
                //            }

                //            List<Element> TempUniqueElementList = new List<Element>();
                //            List<Element> UniqueElementList = new List<Element>();
                //            foreach (ElementLayer el in ElementLayerList.Where(l => l.Layer == Layer))
                //            {
                //                if (dfsuFile.DfsuFileType == DfsuFileType.Dfsu3DSigma || dfsuFile.DfsuFileType == DfsuFileType.Dfsu3DSigmaZ)
                //                {
                //                    if (el.Element.Type == 32)
                //                    {
                //                        bool NodeBigger = false;
                //                        for (int i = 3; i < 6; i++)
                //                        {
                //                            if (el.Element.NodeList[i].Value >= ContourValue)
                //                            {
                //                                NodeBigger = true;
                //                                break;
                //                            }
                //                        }
                //                        if (NodeBigger)
                //                        {
                //                            int countTrue = 0;
                //                            for (int i = 3; i < 6; i++)
                //                            {
                //                                if (el.Element.NodeList[i].Value >= ContourValue && el.Element.NodeList[i].Code == 0)
                //                                {
                //                                    countTrue += 1;
                //                                }
                //                            }
                //                            if (countTrue != el.Element.NodeList.Count)
                //                            {
                //                                TempUniqueElementList.Add(el.Element);
                //                            }
                //                        }
                //                    }
                //                    else if (el.Element.Type == 33)
                //                    {
                //                        bool NodeBigger = false;
                //                        for (int i = 4; i < 8; i++)
                //                        {
                //                            if (el.Element.NodeList[i].Value >= ContourValue)
                //                            {
                //                                NodeBigger = true;
                //                                break;
                //                            }
                //                        }
                //                        if (NodeBigger)
                //                        {
                //                            int countTrue = 0;
                //                            for (int i = 4; i < 8; i++)
                //                            {
                //                                if (el.Element.NodeList[i].Value >= ContourValue && el.Element.NodeList[i].Code == 0)
                //                                {
                //                                    countTrue += 1;
                //                                }
                //                            }
                //                            if (countTrue != el.Element.NodeList.Count)
                //                            {
                //                                TempUniqueElementList.Add(el.Element);
                //                            }
                //                        }
                //                    }
                //                    else
                //                    {
                //                        UpdateTask(AppTaskID, "");
                //                        throw new Exception("Element type is not supported: Element type = [" + el.Element.Type + "]");
                //                    }
                //                }
                //                else
                //                {
                //                    UpdateTask(AppTaskID, "");
                //                    throw new Exception("Bottom only exist for Dfsu3DSigma and Dfsu3DSigmaZ.");
                //                }
                //            }

                //            UniqueElementList = (from el in TempUniqueElementList select el).Distinct().ToList<Element>();

                //            // filling InterpolatedContourNodeList
                //            InterpolatedContourNodeList = new List<Node>();

                //            foreach (Element el in UniqueElementList)
                //            {
                //                if (dfsuFile.DfsuFileType == DfsuFileType.Dfsu3DSigma || dfsuFile.DfsuFileType == DfsuFileType.Dfsu3DSigmaZ)
                //                {
                //                    if (el.Type == 32)
                //                    {
                //                        if (el.NodeList[0].Value >= ContourValue && el.NodeList[1].Value < ContourValue)
                //                        {
                //                            InsertNewNodeInInterpolatedContourNodeList(el.NodeList[0], el.NodeList[1], ContourValue);
                //                        }
                //                        if (el.NodeList[0].Value >= ContourValue && el.NodeList[2].Value < ContourValue)
                //                        {
                //                            InsertNewNodeInInterpolatedContourNodeList(el.NodeList[0], el.NodeList[2], ContourValue);
                //                        }
                //                        if (el.NodeList[1].Value >= ContourValue && el.NodeList[0].Value < ContourValue)
                //                        {
                //                            InsertNewNodeInInterpolatedContourNodeList(el.NodeList[1], el.NodeList[0], ContourValue);
                //                        }
                //                        if (el.NodeList[1].Value >= ContourValue && el.NodeList[2].Value < ContourValue)
                //                        {
                //                            InsertNewNodeInInterpolatedContourNodeList(el.NodeList[1], el.NodeList[2], ContourValue);
                //                        }
                //                        if (el.NodeList[2].Value >= ContourValue && el.NodeList[1].Value < ContourValue)
                //                        {
                //                            InsertNewNodeInInterpolatedContourNodeList(el.NodeList[2], el.NodeList[1], ContourValue);
                //                        }
                //                        if (el.NodeList[2].Value >= ContourValue && el.NodeList[0].Value < ContourValue)
                //                        {
                //                            InsertNewNodeInInterpolatedContourNodeList(el.NodeList[2], el.NodeList[0], ContourValue);
                //                        }
                //                    }
                //                    else if (el.Type == 33)
                //                    {
                //                        if (el.NodeList[0].Value >= ContourValue && el.NodeList[1].Value < ContourValue)
                //                        {
                //                            InsertNewNodeInInterpolatedContourNodeList(el.NodeList[0], el.NodeList[1], ContourValue);
                //                        }
                //                        if (el.NodeList[0].Value >= ContourValue && el.NodeList[3].Value < ContourValue)
                //                        {
                //                            InsertNewNodeInInterpolatedContourNodeList(el.NodeList[0], el.NodeList[3], ContourValue);
                //                        }
                //                        if (el.NodeList[1].Value >= ContourValue && el.NodeList[0].Value < ContourValue)
                //                        {
                //                            InsertNewNodeInInterpolatedContourNodeList(el.NodeList[1], el.NodeList[0], ContourValue);
                //                        }
                //                        if (el.NodeList[1].Value >= ContourValue && el.NodeList[2].Value < ContourValue)
                //                        {
                //                            InsertNewNodeInInterpolatedContourNodeList(el.NodeList[1], el.NodeList[2], ContourValue);
                //                        }
                //                        if (el.NodeList[2].Value >= ContourValue && el.NodeList[1].Value < ContourValue)
                //                        {
                //                            InsertNewNodeInInterpolatedContourNodeList(el.NodeList[2], el.NodeList[1], ContourValue);
                //                        }
                //                        if (el.NodeList[2].Value >= ContourValue && el.NodeList[3].Value < ContourValue)
                //                        {
                //                            InsertNewNodeInInterpolatedContourNodeList(el.NodeList[2], el.NodeList[3], ContourValue);
                //                        }
                //                        if (el.NodeList[3].Value >= ContourValue && el.NodeList[0].Value < ContourValue)
                //                        {
                //                            InsertNewNodeInInterpolatedContourNodeList(el.NodeList[3], el.NodeList[0], ContourValue);
                //                        }
                //                        if (el.NodeList[3].Value >= ContourValue && el.NodeList[2].Value < ContourValue)
                //                        {
                //                            InsertNewNodeInInterpolatedContourNodeList(el.NodeList[3], el.NodeList[2], ContourValue);
                //                        }
                //                    }
                //                    else
                //                    {
                //                        UpdateTask(AppTaskID, "");
                //                        throw new Exception("Element type is not supported: Element type = [" + el.Type + "]");
                //                    }
                //                }
                //                else
                //                {
                //                    UpdateTask(AppTaskID, "");
                //                    throw new Exception("Bottom only exist for Dfsu3DSigma and Dfsu3DSigmaZ.");
                //                }
                //            }

                //            List<Node> UniqueNodeList = (from n in AllNodeList orderby n.ID select n).Distinct().ToList<Node>();

                //            // ------------------------- new code --------------------------
                //            //                     

                //            ForwardVector = new Dictionary<string, Vector>();
                //            BackwardVector = new Dictionary<string, Vector>();

                //            foreach (Element el in UniqueElementList)
                //            {
                //                if (el.Type == 21)
                //                {
                //                    FillVectors21_32(el, UniqueElementList, ContourValue, AppTaskID, false, false);
                //                }
                //                else if (el.Type == 24)
                //                {
                //                    UpdateTask(AppTaskID, "");
                //                    throw new Exception("This should never happen. Node0, Node1 nd Node2 all < ContourValue");
                //                }
                //                else if (el.Type == 25)
                //                {
                //                    FillVectors25_33(el, UniqueElementList, ContourValue, AppTaskID, false, false);
                //                }
                //                else if (el.Type == 32)
                //                {
                //                    FillVectors21_32(el, UniqueElementList, ContourValue, AppTaskID, true, false);
                //                }
                //                else if (el.Type == 33)
                //                {
                //                    FillVectors25_33(el, UniqueElementList, ContourValue, AppTaskID, true, false);
                //                }
                //                else
                //                {
                //                    UpdateTask(AppTaskID, "");
                //                    throw new Exception("Element type is not supported: Element type = [" + el.Type + "]");
                //                }

                //            }

                //            //-------------- new code ------------------------



                //            bool MoreContourLine = true;
                //            while (MoreContourLine)
                //            {
                //                List<Node> FinalContourNodeList = new List<Node>();
                //                Vector LastVector = new Vector();
                //                LastVector = ForwardVector.First().Value;
                //                FinalContourNodeList.Add(LastVector.StartNode);
                //                FinalContourNodeList.Add(LastVector.EndNode);
                //                ForwardVector.Remove(LastVector.StartNode.ID.ToString() + "," + LastVector.EndNode.ID.ToString());
                //                BackwardVector.Remove(LastVector.EndNode.ID.ToString() + "," + LastVector.StartNode.ID.ToString());
                //                bool PolygonCompleted = false;
                //                while (!PolygonCompleted)
                //                {
                //                    List<string> KeyStrList = (from k in ForwardVector.Keys
                //                                               where k.StartsWith(LastVector.EndNode.ID.ToString() + ",")
                //                                               && !k.EndsWith("," + LastVector.StartNode.ID.ToString())
                //                                               select k).ToList<string>();

                //                    if (KeyStrList.Count != 1)
                //                    {
                //                        KeyStrList = (from k in BackwardVector.Keys
                //                                      where k.StartsWith(LastVector.EndNode.ID.ToString() + ",")
                //                                      && !k.EndsWith("," + LastVector.StartNode.ID.ToString())
                //                                      select k).ToList<string>();

                //                        if (KeyStrList.Count != 1)
                //                        {
                //                            PolygonCompleted = true;
                //                            break;
                //                        }
                //                        else
                //                        {
                //                            LastVector = BackwardVector[KeyStrList[0]];
                //                            BackwardVector.Remove(LastVector.StartNode.ID.ToString() + "," + LastVector.EndNode.ID.ToString());
                //                            ForwardVector.Remove(LastVector.EndNode.ID.ToString() + "," + LastVector.StartNode.ID.ToString());
                //                        }
                //                    }
                //                    else
                //                    {
                //                        LastVector = ForwardVector[KeyStrList[0]];
                //                        ForwardVector.Remove(LastVector.StartNode.ID.ToString() + "," + LastVector.EndNode.ID.ToString());
                //                        BackwardVector.Remove(LastVector.EndNode.ID.ToString() + "," + LastVector.StartNode.ID.ToString());
                //                    }
                //                    FinalContourNodeList.Add(LastVector.EndNode);
                //                    if (FinalContourNodeList[FinalContourNodeList.Count - 1] == FinalContourNodeList[0])
                //                    {
                //                        PolygonCompleted = true;
                //                    }
                //                }

                //                if (CalculateAreaOfPolygon(FinalContourNodeList) < 0)
                //                {
                //                    FinalContourNodeList.Reverse();
                //                }

                //                FinalContourNodeList.Add(FinalContourNodeList[0]);
                //                ContourPolygon contourPolygon = new ContourPolygon() { };
                //                contourPolygon.ContourNodeList = FinalContourNodeList;
                //                contourPolygon.ContourValue = ContourValue;
                //                contourPolygon.Layer = Layer;
                //                ContourPolygonList.Add(contourPolygon);

                //                if (ForwardVector.Count == 0)
                //                {
                //                    MoreContourLine = false;
                //                }

                //            }
                //            DrawKMLContourPolygon(ContourPolygonList, dfsuFile, vcount, sbStyleFeacalColiformContour, sbPlacemarkFeacalColiformContour);
                //            vcount += 1;
                //        }
                //        sbPlacemarkFeacalColiformContour.AppendLine(@"</Folder>");
                //        CountContourValue += 1;
                //    }
                //    sbPlacemarkFeacalColiformContour.AppendLine(@"</Folder>");
                //}
                #endregion Bottom of Layer
            }
            sb.AppendLine(@"</Folder>");

            string retStr2 = UpdateAppTaskPercentCompleted(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 100);

            return true;
        }
        private bool WriteKMLMesh(StringBuilder sbHTML, List<ElementLayer> ElementLayerList)
        {
            List<Node> nodeList = new List<Node>();

            sbHTML.AppendLine(@"<Style id=""_line"">");
            sbHTML.AppendLine("<LineStyle>");
            sbHTML.AppendLine(@"<color>ff99ff99</color>");
            sbHTML.AppendLine(@"<width>1</width>");
            sbHTML.AppendLine("</LineStyle>");
            sbHTML.AppendLine(@"</Style>");


            sbHTML.AppendLine(@"<Folder>");
            sbHTML.AppendLine(@"<visibility>0</visibility>");
            sbHTML.AppendLine(@"<name>" + TaskRunnerServiceRes.MIKEMesh + "</name>");

            int CountRefresh = 0;
            int CountAt = 0;
            int UpdateAfter = (int)(ElementLayerList.Count() / 10);
            foreach (ElementLayer ElemLayer in ElementLayerList.Where(c => c.Layer == 1).OrderBy(c => c.Element.ID))
            {
                CountRefresh += 1;
                CountAt += 1;
                if (CountRefresh > UpdateAfter)
                {
                    //_TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, (int)((CountAt * 10) / ElementLayerList.Count()));

                    CountRefresh = 0;
                }

                StringBuilder sbCoord = new StringBuilder();
                float total = 0;
                string LastPart = "";
                foreach (Node node in ElemLayer.Element.NodeList)
                {

                    nodeList.Add(node);

                    if (LastPart == "")
                        LastPart = node.X.ToString().Replace(",", ".") + @"," + node.Y.ToString().Replace(",", ".") + ",0 ";

                    total += node.Z;
                    sbCoord.Append(node.X.ToString().Replace(",", ".") + @"," + node.Y.ToString().Replace(",", ".") + ",0 ");

                }
                sbCoord.Append(LastPart);

                string PolyName = ElemLayer.Element.ID.ToString();

                // Inserting the Placemark
                sbHTML.AppendLine(@"<Placemark>");
                sbHTML.AppendLine(@"<visibility>0</visibility>");
                sbHTML.AppendLine(string.Format(@"<name>{0}</name>", PolyName));
                sbHTML.AppendLine(@"<styleUrl>#_line</styleUrl>");
                sbHTML.AppendLine(@"<LineString>");
                sbHTML.AppendLine(@"<coordinates>");
                sbHTML.AppendLine(sbCoord.ToString());
                sbHTML.AppendLine(@"</coordinates>");
                sbHTML.AppendLine(@"</LineString>");
                sbHTML.AppendLine(@"</Placemark>");
            }

            sbHTML.AppendLine(@"</Folder>");

            //List<Node> uniqueNode = (from n in nodeList
            //                         select n).Distinct().ToList();

            //sbKMLMesh.AppendLine(@"<Folder>");

            //foreach (Node node in uniqueNode)
            //{
            //    sbKMLMesh.AppendLine("<Placemark>");
            //    sbKMLMesh.AppendLine(@"<visibility>0</visibility>");
            //    sbKMLMesh.AppendLine(string.Format(@"<name>{0}</name>", node.ID));
            //    sbKMLMesh.AppendLine(@"<Point>");
            //    sbKMLMesh.AppendLine(string.Format(@"<coordinates>{0},{1},0</coordinates>", node.X, node.Y));
            //    sbKMLMesh.AppendLine("@</Point>");
            //    sbKMLMesh.AppendLine("</Placemark>");
            //}
            //sbKMLMesh.AppendLine(@"</Folder>");

            return true;
        }
        private bool WriteKMLModelInput(StringBuilder sbHTML, List<float> ContourValueList)
        {
            string NotUsed = "";

            PFSFile pfsFile = null;

            if (ContourValueList.Count == 0)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind__, TaskRunnerServiceRes.Parameter, TaskRunnerServiceRes.ContourValues);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotFind__", TaskRunnerServiceRes.Parameter, TaskRunnerServiceRes.ContourValues);
                return false;
            }

            TVFileModel tvFileModelM21_3fm = _TVFileService.GetTVFileModelWithTVItemIDAndTVFileTypeM21FMOrM3FMDB(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID);
            if (!string.IsNullOrWhiteSpace(tvFileModelM21_3fm.Error))
            {
                NotUsed = tvFileModelM21_3fm.Error;
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageList(tvFileModelM21_3fm.Error);
                return false;
            }

            string ServerFileName = tvFileModelM21_3fm.ServerFileName;
            string ServerFilePath = _TVFileService.GetServerFilePath(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID);

            FileInfo fiServer = new FileInfo(ServerFilePath + ServerFileName);

            if (!fiServer.Exists)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFindFile_, fiServer.FullName);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotFindFile_", fiServer.FullName);
                return false;
            }

            pfsFile = new PFSFile(fiServer.FullName);

            MikeScenarioModel mikeScenarioModel = _MikeScenarioService.GetMikeScenarioModelWithMikeScenarioTVItemIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID);
            if (!string.IsNullOrWhiteSpace(mikeScenarioModel.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.MikeScenario, TaskRunnerServiceRes.MikeScenarioTVItemID, _TaskRunnerBaseService._BWObj.appTaskModel.TVItemID.ToString());
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.MikeScenario, TaskRunnerServiceRes.MikeScenarioTVItemID, _TaskRunnerBaseService._BWObj.appTaskModel.TVItemID.ToString());
                return false;
            }

            TVItemModel tvItemModelMikeScenario = _TVItemService.GetTVItemModelWithTVItemIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID);
            if (!string.IsNullOrWhiteSpace(tvItemModelMikeScenario.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, _TaskRunnerBaseService._BWObj.appTaskModel.TVItemID.ToString());
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, _TaskRunnerBaseService._BWObj.appTaskModel.TVItemID.ToString());
                return false;
            }

            List<MikeSourceModel> mikeSourceModelList = _MikeSourceService.GetMikeSourceModelListWithMikeScenarioTVItemIDDB(mikeScenarioModel.MikeScenarioTVItemID);
            if (mikeSourceModelList.Count == 0)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.MikeSource, TaskRunnerServiceRes.MikeScenarioTVItemID, mikeScenarioModel.MikeScenarioTVItemID.ToString());
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.MikeSource, TaskRunnerServiceRes.MikeScenarioTVItemID, mikeScenarioModel.MikeScenarioTVItemID.ToString());
                return false;
            }


            sbHTML.Append("  <Folder><name>" + TaskRunnerServiceRes.ModelInput + "</name>");
            sbHTML.AppendLine(@"  <visibility>0</visibility>");

            #region Source Description
            sbHTML.AppendLine(@"<description><![CDATA[");
            sbHTML.AppendLine(string.Format(@"<h2>{0}</h2>", TaskRunnerServiceRes.ModelParameters));
            sbHTML.AppendLine(@"<ul>");
            sbHTML.AppendLine(string.Format(@"<li><b>{0}:</b> {1:yyyy/MM/dd HH:mm:ss tt}</li>", TaskRunnerServiceRes.ScenarioStartTime, mikeScenarioModel.MikeScenarioStartDateTime_Local));
            sbHTML.AppendLine(string.Format(@"<li><b>{0}:</b> {1:yyyy/MM/dd HH:mm:ss tt}</li>", TaskRunnerServiceRes.ScenarioEndTime, mikeScenarioModel.MikeScenarioEndDateTime_Local));

            foreach (float cv in ContourValueList)
            {
                if (cv >= 14 && cv < 88)
                {
                    sbHTML.AppendLine(string.Format(@"<li><span style=""background-color:Blue; color:White"">{0} = {1:F0}</span</li>", TaskRunnerServiceRes.FCMPNPollutionContour, cv));
                }
                else if (cv >= 88)
                {
                    sbHTML.AppendLine(string.Format(@"<li><span style=""background-color:Red; color:White"">{0} = {1:F0}</span></li>", TaskRunnerServiceRes.FCMPNPollutionContour, cv));
                }
                else
                {
                    sbHTML.AppendLine(string.Format(@"<li><span style=""background-color:Green; color:White"">{0} = {1:F0}</span></li>", TaskRunnerServiceRes.FCMPNPollutionContour, cv));
                }

            }
            sbHTML.AppendLine(@"<li><b>" + TaskRunnerServiceRes.AverageDecayFactor + @":</b> " + mikeScenarioModel.DecayFactor_per_day.ToString("F6").Replace(",", ".") + @" /" + TaskRunnerServiceRes.DayLowerCase + @"</li>");
            if (mikeScenarioModel.DecayIsConstant)
            {
                sbHTML.AppendLine(@"<li><b>" + TaskRunnerServiceRes.DecayIsConstant + @"</b></li>");
            }
            else
            {
                sbHTML.AppendLine(@"<li><b>" + TaskRunnerServiceRes.DecayIsVariable + @"</b></li>");
                sbHTML.AppendLine(@"<ul><li><b>" + TaskRunnerServiceRes.Amplitude + @":</b> " + ((double)mikeScenarioModel.DecayFactorAmplitude).ToString("F6").Replace(",", ".") + @"</li></ul>");
            }
            if (mikeScenarioModel.WindSpeed_km_h > 0)
            {
                sbHTML.AppendLine(@"<li><b>" + TaskRunnerServiceRes.Wind + @":</b> " + mikeScenarioModel.WindSpeed_km_h.ToString("F1").Replace(",", ".") + @" (km/h)   " + (mikeScenarioModel.WindSpeed_km_h / 3.6).ToString("F1").Replace(",", ".") + @" (m/s)</li>");
                sbHTML.AppendLine(@"<li><b>" + TaskRunnerServiceRes.WindDirection + @":</b> " + mikeScenarioModel.WindDirection_deg.ToString("F1").Replace(",", ".") + @" " + TaskRunnerServiceRes.DegreeLowerCase + " (0 = " + TaskRunnerServiceRes.NorthClockwiseLowerCase + @")</li>");
            }
            else
            {
                sbHTML.AppendLine("<li><b>" + TaskRunnerServiceRes.NoWind + @"</b></li>");
            }
            sbHTML.AppendLine(@"<li><b>" + TaskRunnerServiceRes.Temperature + @":</b> " + mikeScenarioModel.AmbientTemperature_C.ToString("F1").Replace(",", ".") + " " + TaskRunnerServiceRes.Celcius + @"</li>");
            sbHTML.AppendLine(@"<li><b>" + TaskRunnerServiceRes.Salinity + @":</b> " + mikeScenarioModel.AmbientSalinity_PSU.ToString("F1").Replace(",", ".") + " " + TaskRunnerServiceRes.PSU + @"</li>");
            sbHTML.AppendLine(@"<li><b>" + TaskRunnerServiceRes.ManningNumber + @":</b> " + mikeScenarioModel.ManningNumber.ToString().Replace(",", ".") + @"</li>");
            sbHTML.AppendLine(string.Format(@"<li><b>" + TaskRunnerServiceRes.ResultFrequency + @":</b> {0:F0} {1}</li>", mikeScenarioModel.ResultFrequency_min, TaskRunnerServiceRes.MinutesLowerCase));
            sbHTML.AppendLine(@"</ul>");

            List<MikeSourceModel> mikeSourceModelListAll = mikeSourceModelList.Where(c => c.Include == true && c.IsRiver == false).Concat(mikeSourceModelList.Where(c => c.Include == false && c.IsRiver == false).Concat(mikeSourceModelList.Where(c => c.IsRiver == true))).ToList();

            // Do Mike Source 
            foreach (MikeSourceModel mikeSourceModel in mikeSourceModelListAll)
            {
                List<MikeSourceStartEndModel> mikeSourceStartEndModelListLocal = _MikeSourceStartEndService.GetMikeSourceStartEndModelListWithMikeSourceIDDB(mikeSourceModel.MikeSourceID);

                if (mikeSourceStartEndModelListLocal.Count == 0)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.MikeSourceStartEnd, TaskRunnerServiceRes.MikeSourceID, mikeSourceModel.MikeSourceID.ToString());
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.MikeSourceStartEnd, TaskRunnerServiceRes.MikeSourceID, mikeSourceModel.MikeSourceID.ToString());
                }

                if (mikeSourceModel.Include)
                {
                    if (mikeSourceModel.IsRiver)
                    {
                        sbHTML.AppendLine(string.Format(@"<h2 style='Color: Blue'>{0} ({1})</h2>", mikeSourceModel.MikeSourceTVText, TaskRunnerServiceRes.IncludedLowerCase));
                    }
                    else
                    {
                        sbHTML.AppendLine(string.Format(@"<h2 style='Color: Green'>{0} ({1})</h2>", mikeSourceModel.MikeSourceTVText, TaskRunnerServiceRes.IncludedLowerCase));
                    }
                }
                else
                {
                    sbHTML.AppendLine(string.Format(@"<h2 style='Color: Red'>{0} ({1})</h2>", mikeSourceModel.MikeSourceTVText, TaskRunnerServiceRes.NotIncludedLowerCase));
                }
                sbHTML.AppendLine(@"<h3>" + TaskRunnerServiceRes.Effluent + @"</h3>");
                sbHTML.AppendLine(@"<ul>");
                sbHTML.AppendLine(@"<li><b>" + TaskRunnerServiceRes.Coordinates + @"</b> " + string.Format(@"&nbsp;&nbsp;&nbsp; {0:F5} &nbsp; {1:F5}</li>", mikeSourceModel.Lat, mikeSourceModel.Lng).Replace(",", "."));

                if ((bool)mikeSourceModel.IsContinuous)
                {
                    sbHTML.AppendLine(@"<li><b>" + TaskRunnerServiceRes.IsContinuous + @"</b></li>");
                    sbHTML.AppendLine(@"<li><b>" + TaskRunnerServiceRes.Flow + @":</b> " + (mikeSourceStartEndModelListLocal[0].SourceFlowStart_m3_day / 24 / 3600).ToString("F6").Replace(",", ".") + " (m3/s)  " + mikeSourceStartEndModelListLocal[0].SourceFlowStart_m3_day.ToString("F1").Replace(",", ".") + @" (m3/" + TaskRunnerServiceRes.DayLowerCase + @")</li>");
                    sbHTML.AppendLine(@"<li><b>" + TaskRunnerServiceRes.FCMPNPer100ML + @":</b> " + mikeSourceStartEndModelListLocal[0].SourcePollutionStart_MPN_100ml.ToString("F0").Replace(",", ".") + @"</li>");
                    sbHTML.AppendLine(@"<li><b>" + TaskRunnerServiceRes.Temperature + @":</b> " + mikeSourceStartEndModelListLocal[0].SourceTemperatureStart_C.ToString("F1").Replace(",", ".") + @" " + TaskRunnerServiceRes.Celcius + @"</li>");
                    sbHTML.AppendLine(@"<li><b>" + TaskRunnerServiceRes.Salinity + @":</b> " + mikeSourceStartEndModelListLocal[0].SourceSalinityStart_PSU.ToString("F1").Replace(",", ".") + @" " + TaskRunnerServiceRes.PSU + @"</li>");
                }
                else
                {
                    sbHTML.AppendLine(@"<li><b>" + TaskRunnerServiceRes.IsNotContinuous + @"</b></li>");

                    int CountMikeSourceStartEnd = 0;
                    foreach (MikeSourceStartEndModel mssem in mikeSourceStartEndModelListLocal)
                    {
                        CountMikeSourceStartEnd += 1;
                        sbHTML.AppendLine(@"<br /><b>" + TaskRunnerServiceRes.Spill + @": " + CountMikeSourceStartEnd + "</b><br />");
                        sbHTML.AppendLine(string.Format(@"<li><b>" + TaskRunnerServiceRes.SpillStartTime + @":</b> {0:yyyy/MM/dd HH:mm:ss tt} (UTC)</li>", mssem.StartDateAndTime_Local));
                        sbHTML.AppendLine(string.Format(@"<li><b>" + TaskRunnerServiceRes.SpillEndTime + @":</b> {0:yyyy/MM/dd HH:mm:ss tt} (UTC)</li>", mssem.EndDateAndTime_Local));
                        sbHTML.AppendLine(@"<li><b>" + TaskRunnerServiceRes.FlowStart + @":</b> " + ((double)mssem.SourceFlowStart_m3_day / 24 / 3600).ToString("F6").Replace(",", ".") + @" (m3/s)  " + ((double)mssem.SourceFlowStart_m3_day).ToString("F0").Replace(",", ".") + @" (m3/" + TaskRunnerServiceRes.DayLowerCase + @")</li>");
                        sbHTML.AppendLine(@"<li><b>" + TaskRunnerServiceRes.FlowEnd + @":</b> " + ((double)mssem.SourceFlowEnd_m3_day / 24 / 3600).ToString("F6").Replace(",", ".") + @" (m3/s)  " + ((double)mssem.SourceFlowEnd_m3_day).ToString("F0").Replace(",", ".") + @" (m3/" + TaskRunnerServiceRes.DayLowerCase + @")</li>");
                        sbHTML.AppendLine(@"<li><b>" + TaskRunnerServiceRes.FCMPNPer100MLStart + @":</b> " + ((double)mssem.SourcePollutionStart_MPN_100ml).ToString("F0").Replace(",", ".") + @"</li>");
                        sbHTML.AppendLine(@"<li><b>" + TaskRunnerServiceRes.FCMPNPer100MLEnd + @":</b> " + ((double)mssem.SourcePollutionEnd_MPN_100ml).ToString("F0").Replace(",", ".") + @"</li>");
                        sbHTML.AppendLine(@"<li><b>" + TaskRunnerServiceRes.TemperatureStart + @":</b> " + ((double)mssem.SourceTemperatureStart_C).ToString("F0").Replace(",", ".") + @" " + TaskRunnerServiceRes.Celcius + @"</li>");
                        sbHTML.AppendLine(@"<li><b>" + TaskRunnerServiceRes.TemperatureEnd + @":</b> " + ((double)mssem.SourceTemperatureEnd_C).ToString("F0").Replace(",", ".") + @" " + TaskRunnerServiceRes.Celcius + @"</li>");
                        sbHTML.AppendLine(@"<li><b>" + TaskRunnerServiceRes.SalinityStart + @":</b> " + ((double)mssem.SourceSalinityStart_PSU).ToString("F0").Replace(",", ".") + @" " + TaskRunnerServiceRes.PSU + @"</li>");
                        sbHTML.AppendLine(@"<li><b>" + TaskRunnerServiceRes.SalinityEnd + @":</b> " + ((double)mssem.SourceSalinityEnd_PSU).ToString("F0").Replace(",", ".") + @" " + TaskRunnerServiceRes.PSU + @"</li>");
                    }
                }
                sbHTML.AppendLine(@"</ul>");
            }


            sbHTML.AppendLine(@"<iframe src=""about:"" width=""600"" height=""1"" />");
            sbHTML.AppendLine(@"]]></description>");

            #endregion Source Description

            sbHTML.Append(" <Folder>");
            sbHTML.Append("    <name>" + TaskRunnerServiceRes.SourceIncluded + @"</name>");
            sbHTML.AppendLine(@"    <visibility>0</visibility>");

            for (int i = 1; i < 1000; i++)
            {
                PFSSection pfsSectionSource = pfsFile.GetSectionFromHandle("FemEngineHD/HYDRODYNAMIC_MODULE/SOURCES/SOURCE_" + i.ToString());

                if (pfsSectionSource == null)
                {
                    break;
                }

                int? SourceIncluded = GetSourceIncluded(pfsSectionSource);
                if (SourceIncluded == null)
                {
                    pfsFile.Close();
                    return false;
                }

                if (SourceIncluded == 1)
                {
                    MikeSourceModel mikeSourceModelLocal = (from msl in mikeSourceModelList
                                                            where msl.SourceNumberString == "SOURCE_" + i.ToString()
                                                            select msl).FirstOrDefault<MikeSourceModel>();

                    if (mikeSourceModelLocal == null)
                    {
                        NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.MikeSource, TaskRunnerServiceRes.SourceNumberString, "Source_" + i.ToString());
                        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.MikeSource, TaskRunnerServiceRes.SourceNumberString, "Source_" + i.ToString());
                        pfsFile.Close();
                        return false;
                    }

                    if (!WriteKMLSourcePlacemark(sbHTML, pfsFile, pfsSectionSource, i, mikeSourceModelLocal))
                    {
                        pfsFile.Close();
                        return false;
                    }
                }

            }

            sbHTML.Append("</Folder>");

            sbHTML.Append("<Folder><name>" + TaskRunnerServiceRes.SourceNotIncluded + @"</name>");
            sbHTML.AppendLine(@"<visibility>0</visibility>");

            // showing not used sources 
            for (int i = 1; i < 1000; i++)
            {
                PFSSection pfsSectionSource = pfsFile.GetSectionFromHandle("FemEngineHD/HYDRODYNAMIC_MODULE/SOURCES/SOURCE_" + i.ToString());

                if (pfsSectionSource == null)
                {
                    break;
                }

                int? SourceIncluded = GetSourceIncluded(pfsSectionSource);
                if (SourceIncluded == null)
                {
                    pfsFile.Close();
                    return false;
                }

                if (SourceIncluded == 0)
                {
                    MikeSourceModel mikeSourceModelLocal = (from msl in mikeSourceModelList
                                                            where msl.SourceNumberString == "SOURCE_" + i.ToString()
                                                            select msl).FirstOrDefault<MikeSourceModel>();

                    if (mikeSourceModelLocal == null)
                    {
                        NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.MikeSource, TaskRunnerServiceRes.SourceNumberString, "Source_" + i.ToString());
                        _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.MikeSource, TaskRunnerServiceRes.SourceNumberString, "Source_" + i.ToString());
                        pfsFile.Close();
                        return false;
                    }

                    if (!WriteKMLSourcePlacemark(sbHTML, pfsFile, pfsSectionSource, i, mikeSourceModelLocal))
                    {
                        pfsFile.Close();
                        return false;
                    }
                }

            }

            sbHTML.Append("</Folder>");

            sbHTML.Append("</Folder>");

            return true;
        }
        private bool WriteKMLPollutionLimitsContourLine(DfsuFile dfsuFile, StringBuilder sbHTML, List<ElementLayer> elementLayerList, List<NodeLayer> topNodeLayerList, List<NodeLayer> bottomNodeLayerList, List<float> ContourValueList)
        {
            string NotUsed = "";

            List<List<ContourPolygon>> ContourPolygonListList = new List<List<ContourPolygon>>();
            List<int> LayerList = new List<int>();

            int PercentCompleted = 3;
            int ItemNumber = 0;

            // getting the ItemNumber
            foreach (IDfsSimpleDynamicItemInfo dfsDyInfo in dfsuFile.ItemInfo)
            {
                if (dfsDyInfo.Quantity.Item == eumItem.eumIConcentration
                    || dfsDyInfo.Quantity.Item == eumItem.eumIConcentration1
                    || dfsDyInfo.Quantity.Item == eumItem.eumIConcentration_1
                    || dfsDyInfo.Quantity.Item == eumItem.eumIConcentration_2
                    || dfsDyInfo.Quantity.Item == eumItem.eumIConcentration_3
                    || dfsDyInfo.Quantity.Item == eumItem.eumIConcentration_4)
                {
                    ItemNumber = dfsDyInfo.ItemNumber;
                    break;
                }
            }

            if (ItemNumber == 0)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind__, TaskRunnerServiceRes.ParameterType, "eumIConcentration1");
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotFind__", TaskRunnerServiceRes.ParameterType, "eumIConcentration1");
                return false;
            }

            //int pcount = 0;
            int CountLayer = (dfsuFile.NumberOfSigmaLayers == 0 ? 1 : dfsuFile.NumberOfSigmaLayers);
            int CountAt = 0;

            int TotalCount = CountLayer * ContourValueList.Count;

            for (int Layer = 1; Layer <= 1 /* CountLayer */; Layer++)
            {
                if (!LayerList.Contains(Layer))
                {
                    LayerList.Add(Layer);
                }

                #region Top of Layer
                int CountContour = 1;
                foreach (float ContourValue in ContourValueList)
                {
                    CountAt += 1;

                    int PercentCompletedTemp = ((int)(((CountAt - 1) * (float)100.0) / (float)TotalCount));
                    if (PercentCompleted > 3 && PercentCompleted != PercentCompletedTemp)
                    {
                        string retStr = UpdateAppTaskPercentCompleted(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, PercentCompleted);
                    }
                    PercentCompleted = PercentCompletedTemp;

                    List<Node> AllNodeList = new List<Node>();
                    List<ContourPolygon> ContourPolygonList = new List<ContourPolygon>();

                    for (int timeStep = 0; timeStep < dfsuFile.NumberOfTimeSteps; timeStep++)
                    {
                        float[] ValueList = (float[])dfsuFile.ReadItemTimeStep(ItemNumber, timeStep).Data;

                        for (int i = 0; i < elementLayerList.Count; i++)
                        {
                            if (elementLayerList[i].Element.Value < ValueList[i])
                            {
                                elementLayerList[i].Element.Value = ValueList[i];
                            }
                        }
                    }

                    List<ElementLayer> elementLayerListAtLayer = elementLayerList.Where(c => c.Layer == Layer).ToList();
                    List<NodeLayer> topNodeLayerListAtLayer = topNodeLayerList.Where(c => c.Layer == Layer).ToList();

                    foreach (NodeLayer nl in topNodeLayerListAtLayer)
                    {
                        float Total = 0;
                        foreach (Element element in nl.Node.ElementList)
                        {
                            Total += element.Value;
                        }
                        nl.Node.Value = Total / nl.Node.ElementList.Count;
                    }


                    List<NodeLayer> AboveNodeLayerList = (from n in topNodeLayerListAtLayer
                                                          where (n.Node.Value >= ContourValue)
                                                          && n.Layer == Layer
                                                          select n).ToList<NodeLayer>();

                    foreach (NodeLayer snl in AboveNodeLayerList)
                    {
                        List<NodeLayer> EndNodeLayerList = null;

                        List<NodeLayer> NodeLayerConnectedList = (from nll in topNodeLayerListAtLayer
                                                                  from n in snl.Node.ConnectNodeList
                                                                  where (n.ID == nll.Node.ID)
                                                                  select nll).ToList<NodeLayer>();

                        EndNodeLayerList = (from nll in NodeLayerConnectedList
                                            where (nll.Node.ID != snl.Node.ID)
                                            && (nll.Node.Value < ContourValue)
                                            && nll.Layer == Layer
                                            select nll).ToList<NodeLayer>();

                        foreach (NodeLayer en in EndNodeLayerList)
                        {
                            AllNodeList.Add(en.Node);
                        }

                        if (snl.Node.Code != 0)
                        {
                            AllNodeList.Add(snl.Node);
                        }

                    }

                    List<Element> TempUniqueElementList = new List<Element>();
                    List<Element> UniqueElementList = new List<Element>();
                    foreach (ElementLayer el in elementLayerListAtLayer)
                    {
                        if (dfsuFile.DfsuFileType == DfsuFileType.Dfsu3DSigma || dfsuFile.DfsuFileType == DfsuFileType.Dfsu3DSigmaZ)
                        {
                            if (el.Element.Type == 32)
                            {
                                bool NodeBigger = false;
                                for (int i = 3; i < 6; i++)
                                {
                                    if (el.Element.NodeList[i].Value >= ContourValue)
                                    {
                                        NodeBigger = true;
                                        break;
                                    }
                                }
                                if (NodeBigger)
                                {
                                    int countTrue = 0;
                                    for (int i = 3; i < 6; i++)
                                    {
                                        if (el.Element.NodeList[i].Value >= ContourValue && el.Element.NodeList[i].Code == 0)
                                        {
                                            countTrue += 1;
                                        }
                                    }
                                    if (countTrue != el.Element.NodeList.Count)
                                    {
                                        TempUniqueElementList.Add(el.Element);
                                    }
                                }
                            }
                            else if (el.Element.Type == 33)
                            {
                                bool NodeBigger = false;
                                for (int i = 4; i < 8; i++)
                                {
                                    if (el.Element.NodeList[i].Value >= ContourValue)
                                    {
                                        NodeBigger = true;
                                        break;
                                    }
                                }
                                if (NodeBigger)
                                {
                                    int countTrue = 0;
                                    for (int i = 4; i < 8; i++)
                                    {
                                        if (el.Element.NodeList[i].Value >= ContourValue && el.Element.NodeList[i].Code == 0)
                                        {
                                            countTrue += 1;
                                        }
                                    }
                                    if (countTrue != el.Element.NodeList.Count)
                                    {
                                        TempUniqueElementList.Add(el.Element);
                                    }
                                }
                            }
                            else
                            {
                                NotUsed = string.Format(TaskRunnerServiceRes.ElementType_IsNotSupported, el.Element.Type.ToString());
                                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("ElementType_IsNotSupported", el.Element.Type.ToString());
                                return false;
                            }
                        }
                        else if (dfsuFile.DfsuFileType == DfsuFileType.Dfsu2D)
                        {
                            bool NodeBigger = false;
                            for (int i = 0; i < el.Element.NodeList.Count; i++)
                            {
                                if (el.Element.NodeList[i].Value >= ContourValue)
                                {
                                    NodeBigger = true;
                                    break;
                                }
                            }
                            if (NodeBigger)
                            {
                                int countTrue = 0;
                                for (int i = 0; i < el.Element.NodeList.Count; i++)
                                {
                                    if (el.Element.NodeList[i].Value >= ContourValue && el.Element.NodeList[i].Code == 0)
                                    {
                                        countTrue += 1;
                                    }
                                }
                                if (countTrue != el.Element.NodeList.Count)
                                {
                                    TempUniqueElementList.Add(el.Element);
                                }
                            }
                        }
                    }

                    UniqueElementList = (from el in TempUniqueElementList select el).Distinct().ToList<Element>();

                    // filling InterpolatedContourNodeList
                    InterpolatedContourNodeList = new List<Node>();

                    int count = 0;
                    foreach (Element el in UniqueElementList)
                    {
                        count += 1;
                        if (dfsuFile.DfsuFileType == DfsuFileType.Dfsu3DSigma || dfsuFile.DfsuFileType == DfsuFileType.Dfsu3DSigmaZ)
                        {
                            if (el.Type == 32)
                            {
                                if (el.NodeList[3].Value >= ContourValue && el.NodeList[4].Value < ContourValue)
                                {
                                    InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[3], el.NodeList[4], ContourValue);
                                }
                                if (el.NodeList[3].Value >= ContourValue && el.NodeList[5].Value < ContourValue)
                                {
                                    InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[3], el.NodeList[5], ContourValue);
                                }
                                if (el.NodeList[4].Value >= ContourValue && el.NodeList[3].Value < ContourValue)
                                {
                                    InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[4], el.NodeList[3], ContourValue);
                                }
                                if (el.NodeList[4].Value >= ContourValue && el.NodeList[5].Value < ContourValue)
                                {
                                    InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[4], el.NodeList[5], ContourValue);
                                }
                                if (el.NodeList[5].Value >= ContourValue && el.NodeList[4].Value < ContourValue)
                                {
                                    InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[5], el.NodeList[4], ContourValue);
                                }
                                if (el.NodeList[5].Value >= ContourValue && el.NodeList[3].Value < ContourValue)
                                {
                                    InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[5], el.NodeList[3], ContourValue);
                                }
                            }
                            else if (el.Type == 33)
                            {
                                if (el.NodeList[4].Value >= ContourValue && el.NodeList[5].Value < ContourValue)
                                {
                                    InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[4], el.NodeList[5], ContourValue);
                                }
                                if (el.NodeList[4].Value >= ContourValue && el.NodeList[7].Value < ContourValue)
                                {
                                    InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[4], el.NodeList[7], ContourValue);
                                }
                                if (el.NodeList[5].Value >= ContourValue && el.NodeList[4].Value < ContourValue)
                                {
                                    InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[5], el.NodeList[4], ContourValue);
                                }
                                if (el.NodeList[5].Value >= ContourValue && el.NodeList[6].Value < ContourValue)
                                {
                                    InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[5], el.NodeList[6], ContourValue);
                                }
                                if (el.NodeList[6].Value >= ContourValue && el.NodeList[5].Value < ContourValue)
                                {
                                    InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[6], el.NodeList[5], ContourValue);
                                }
                                if (el.NodeList[6].Value >= ContourValue && el.NodeList[7].Value < ContourValue)
                                {
                                    InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[6], el.NodeList[7], ContourValue);
                                }
                                if (el.NodeList[7].Value >= ContourValue && el.NodeList[4].Value < ContourValue)
                                {
                                    InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[7], el.NodeList[4], ContourValue);
                                }
                                if (el.NodeList[7].Value >= ContourValue && el.NodeList[6].Value < ContourValue)
                                {
                                    InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[7], el.NodeList[6], ContourValue);
                                }
                            }
                            else
                            {
                                NotUsed = string.Format(TaskRunnerServiceRes.ElementType_IsNotSupported, el.Type.ToString());
                                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("ElementType_IsNotSupported", el.Type.ToString());
                                return false;
                            }
                        }
                        else if (dfsuFile.DfsuFileType == DfsuFileType.Dfsu2D)
                        {
                            if (el.Type == 21)
                            {
                                if (el.NodeList[0].Value >= ContourValue && el.NodeList[1].Value < ContourValue)
                                {
                                    InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[0], el.NodeList[1], ContourValue);
                                }
                                if (el.NodeList[0].Value >= ContourValue && el.NodeList[2].Value < ContourValue)
                                {
                                    InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[0], el.NodeList[2], ContourValue);
                                }
                                if (el.NodeList[1].Value >= ContourValue && el.NodeList[0].Value < ContourValue)
                                {
                                    InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[1], el.NodeList[0], ContourValue);
                                }
                                if (el.NodeList[1].Value >= ContourValue && el.NodeList[2].Value < ContourValue)
                                {
                                    InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[1], el.NodeList[2], ContourValue);
                                }
                                if (el.NodeList[2].Value >= ContourValue && el.NodeList[1].Value < ContourValue)
                                {
                                    InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[2], el.NodeList[1], ContourValue);
                                }
                                if (el.NodeList[2].Value >= ContourValue && el.NodeList[0].Value < ContourValue)
                                {
                                    InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[2], el.NodeList[0], ContourValue);
                                }
                            }
                            else if (el.Type == 24)
                            {
                            }
                            else if (el.Type == 25)
                            {
                                if (el.NodeList[0].Value >= ContourValue && el.NodeList[1].Value < ContourValue)
                                {
                                    InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[0], el.NodeList[1], ContourValue);
                                }
                                if (el.NodeList[0].Value >= ContourValue && el.NodeList[3].Value < ContourValue)
                                {
                                    InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[0], el.NodeList[3], ContourValue);
                                }
                                if (el.NodeList[1].Value >= ContourValue && el.NodeList[0].Value < ContourValue)
                                {
                                    InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[1], el.NodeList[0], ContourValue);
                                }
                                if (el.NodeList[1].Value >= ContourValue && el.NodeList[2].Value < ContourValue)
                                {
                                    InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[1], el.NodeList[2], ContourValue);
                                }
                                if (el.NodeList[2].Value >= ContourValue && el.NodeList[1].Value < ContourValue)
                                {
                                    InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[2], el.NodeList[1], ContourValue);
                                }
                                if (el.NodeList[2].Value >= ContourValue && el.NodeList[3].Value < ContourValue)
                                {
                                    InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[2], el.NodeList[3], ContourValue);
                                }
                                if (el.NodeList[3].Value >= ContourValue && el.NodeList[0].Value < ContourValue)
                                {
                                    InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[3], el.NodeList[0], ContourValue);
                                }
                                if (el.NodeList[3].Value >= ContourValue && el.NodeList[2].Value < ContourValue)
                                {
                                    InsertNewNodeInInterpolatedContourNodeList(dfsuFile, el.NodeList[3], el.NodeList[2], ContourValue);
                                }
                            }
                            else
                            {
                                NotUsed = string.Format(TaskRunnerServiceRes.ElementType_IsNotSupported, el.Type.ToString());
                                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("ElementType_IsNotSupported", el.Type.ToString());
                                return false;
                            }
                        }
                    }
                    List<Node> UniqueNodeList = (from n in AllNodeList orderby n.ID select n).Distinct().ToList<Node>();

                    ForwardVector = new Dictionary<String, Vector>();
                    BackwardVector = new Dictionary<String, Vector>();

                    // ------------------------- new code --------------------------
                    //                     

                    foreach (Element el in UniqueElementList)
                    {
                        if (el.Type == 21)
                        {
                            FillVectors21_32(el, UniqueElementList, ContourValue, false, true);
                        }
                        else if (el.Type == 24)
                        {
                            NotUsed = string.Format(TaskRunnerServiceRes.ElementType_IsNotSupported, el.Type.ToString());
                            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("ElementType_IsNotSupported", el.Type.ToString());
                            return false;
                        }
                        else if (el.Type == 25)
                        {
                            FillVectors25_33(el, UniqueElementList, ContourValue, false, true);
                        }
                        else if (el.Type == 32)
                        {
                            FillVectors21_32(el, UniqueElementList, ContourValue, true, true);
                        }
                        else if (el.Type == 33)
                        {
                            FillVectors25_33(el, UniqueElementList, ContourValue, true, true);
                        }
                        else
                        {
                            NotUsed = string.Format(TaskRunnerServiceRes.ElementType_IsNotSupported, el.Type.ToString());
                            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("ElementType_IsNotSupported", el.Type.ToString());
                            return false;
                        }

                    }

                    bool MoreContourLine = true;
                    while (MoreContourLine && ForwardVector.Count > 0)
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

                        if (_MapInfoService.CalculateAreaOfPolygon(FinalContourNodeList) < 0)
                        {
                            FinalContourNodeList.Reverse();
                        }

                        FinalContourNodeList.Add(FinalContourNodeList[0]);
                        ContourPolygon contourPolygon = new ContourPolygon() { };
                        contourPolygon.ContourNodeList = FinalContourNodeList;
                        contourPolygon.ContourValue = ContourValue;
                        contourPolygon.Layer = Layer;

                        ContourPolygonList.Add(contourPolygon);

                        if (ForwardVector.Count == 0)
                        {
                            MoreContourLine = false;
                        }
                    }

                    ContourPolygonListList.Add(ContourPolygonList);

                    CountContour += 1;
                }
                #endregion Top of Layer

                #region Bottom of Layer
                //// 
                //if (Layer == dfsuFile.NumberOfSigmaLayers)
                //{
                //    sbKMLPollutionLimitsContour.AppendLine(string.Format(@"<Folder><name>Bottom of Layer [{0}]</name>", Layer));
                //    sbKMLPollutionLimitsContour.AppendLine(@"<visibility>0</visibility>");
                //    CountContour = 1;
                //    foreach (float ContourValue in ContourValueList)
                //    {
                //        CountAt += 1;
                //        sbKMLPollutionLimitsContour.AppendLine(string.Format(@"<Folder><name>Contour Value [{0}]</name>", ContourValue));
                //        sbKMLPollutionLimitsContour.AppendLine(@"<visibility>0</visibility>");
                //        string AppTaskStatus = ((int)((CountAt * 100) / (SigmaLayerValueList.Count * ContourValueList.Count))).ToString() + " %";
                //        UpdateTask(AppTaskID, AppTaskStatus);

                //        List<Node> AllNodeList = new List<Node>();
                //        List<ContourPolygon> ContourPolygonList = new List<ContourPolygon>();

                //        //foreach (Dfs.Parameter.TimeSeriesValue v in p.TimeSeriesValueList)
                //        //{
                //        for (int timeStep = 0; timeStep < dfsuFile.NumberOfTimeSteps; timeStep++)
                //        {

                //            float[] ValueList = (float[])dfsuFile.ReadItemTimeStep(ItemNumber, timeStep).Data;

                //            for (int i = 0; i < ElementLayerList.Count; i++)
                //            {
                //                if (ElementLayerList[i].Element.Value < ValueList[i])
                //                {
                //                    ElementLayerList[i].Element.Value = ValueList[i];
                //                }
                //            }
                //        }
                //        //}

                //        foreach (NodeLayer nl in BottomNodeLayerList)
                //        {
                //            float Total = 0;
                //            foreach (Element element in nl.Node.ElementList)
                //            {
                //                Total += element.Value;
                //            }
                //            nl.Node.Value = Total / nl.Node.ElementList.Count;
                //        }

                //        List<NodeLayer> AboveNodeLayerList = new List<NodeLayer>();

                //        AboveNodeLayerList = (from n in BottomNodeLayerList
                //                              where (n.Node.Value >= ContourValue)
                //                              && n.Layer == Layer
                //                              select n).ToList<NodeLayer>();

                //        foreach (NodeLayer snl in AboveNodeLayerList)
                //        {
                //            List<NodeLayer> EndNodeLayerList = null;

                //            List<NodeLayer> NodeLayerConnectedList = (from nll in BottomNodeLayerList
                //                                                      from n in snl.Node.ConnectNodeList
                //                                                      where (n.ID == nll.Node.ID)
                //                                                      select nll).ToList<NodeLayer>();

                //            EndNodeLayerList = (from nll in NodeLayerConnectedList
                //                                where (nll.Node.ID != snl.Node.ID)
                //                                && (nll.Node.Value < ContourValue)
                //                                && nll.Layer == Layer
                //                                select nll).ToList<NodeLayer>();

                //            foreach (NodeLayer en in EndNodeLayerList)
                //            {
                //                AllNodeList.Add(en.Node);
                //            }

                //            if (snl.Node.Code != 0)
                //            {
                //                AllNodeList.Add(snl.Node);
                //            }

                //        }

                //        if (AllNodeList.Count == 0)
                //        {
                //            continue;
                //        }

                //        List<Element> TempUniqueElementList = new List<Element>();
                //        List<Element> UniqueElementList = new List<Element>();
                //        foreach (ElementLayer el in ElementLayerList.Where(l => l.Layer == Layer))
                //        {
                //            if (dfsuFile.DfsuFileType == DfsuFileType.Dfsu3DSigma || dfsuFile.DfsuFileType == DfsuFileType.Dfsu3DSigmaZ)
                //            {
                //                if (el.Element.Type == 32)
                //                {
                //                    bool NodeBigger = false;
                //                    for (int i = 0; i < 3; i++)
                //                    {
                //                        if (el.Element.NodeList[i].Value >= ContourValue)
                //                        {
                //                            NodeBigger = true;
                //                            break;
                //                        }
                //                    }
                //                    if (NodeBigger)
                //                    {
                //                        int countTrue = 0;
                //                        for (int i = 0; i < 3; i++)
                //                        {
                //                            if (el.Element.NodeList[i].Value >= ContourValue && el.Element.NodeList[i].Code == 0)
                //                            {
                //                                countTrue += 1;
                //                            }
                //                        }
                //                        if (countTrue != el.Element.NodeList.Count)
                //                        {
                //                            TempUniqueElementList.Add(el.Element);
                //                        }
                //                    }
                //                }
                //                else if (el.Element.Type == 33)
                //                {
                //                    bool NodeBigger = false;
                //                    for (int i = 0; i < 4; i++)
                //                    {
                //                        if (el.Element.NodeList[i].Value >= ContourValue)
                //                        {
                //                            NodeBigger = true;
                //                            break;
                //                        }
                //                    }
                //                    if (NodeBigger)
                //                    {
                //                        int countTrue = 0;
                //                        for (int i = 0; i < 4; i++)
                //                        {
                //                            if (el.Element.NodeList[i].Value >= ContourValue && el.Element.NodeList[i].Code == 0)
                //                            {
                //                                countTrue += 1;
                //                            }
                //                        }
                //                        if (countTrue != el.Element.NodeList.Count)
                //                        {
                //                            TempUniqueElementList.Add(el.Element);
                //                        }
                //                    }
                //                }
                //                else
                //                {
                //                    UpdateTask(AppTaskID, "");
                //                    throw new Exception("Element type is not supported: Element type = [" + el.Element.Type + "]");
                //                }
                //            }
                //            else if (dfsuFile.DfsuFileType == DfsuFileType.Dfsu2D)
                //            {
                //                bool NodeBigger = false;
                //                for (int i = 0; i < el.Element.NodeList.Count; i++)
                //                {
                //                    if (el.Element.NodeList[i].Value >= ContourValue)
                //                    {
                //                        NodeBigger = true;
                //                        break;
                //                    }
                //                }
                //                if (NodeBigger)
                //                {
                //                    int countTrue = 0;
                //                    for (int i = 0; i < el.Element.NodeList.Count; i++)
                //                    {
                //                        if (el.Element.NodeList[i].Value >= ContourValue && el.Element.NodeList[i].Code == 0)
                //                        {
                //                            countTrue += 1;
                //                        }
                //                    }
                //                    if (countTrue != el.Element.NodeList.Count)
                //                    {
                //                        TempUniqueElementList.Add(el.Element);
                //                    }
                //                }
                //            }
                //        }

                //        UniqueElementList = (from el in TempUniqueElementList select el).Distinct().ToList<Element>();

                //        // filling InterpolatedContourNodeList
                //        InterpolatedContourNodeList = new List<Node>();

                //        foreach (Element el in UniqueElementList)
                //        {
                //            if (dfsuFile.DfsuFileType == DfsuFileType.Dfsu3DSigma || dfsuFile.DfsuFileType == DfsuFileType.Dfsu3DSigmaZ)
                //            {
                //                if (el.Type == 32)
                //                {
                //                    if (el.NodeList[0].Value >= ContourValue && el.NodeList[1].Value < ContourValue)
                //                    {
                //                        InsertNewNodeInInterpolatedContourNodeList(el.NodeList[0], el.NodeList[1], ContourValue);
                //                    }
                //                    if (el.NodeList[0].Value >= ContourValue && el.NodeList[2].Value < ContourValue)
                //                    {
                //                        InsertNewNodeInInterpolatedContourNodeList(el.NodeList[0], el.NodeList[2], ContourValue);
                //                    }
                //                    if (el.NodeList[1].Value >= ContourValue && el.NodeList[0].Value < ContourValue)
                //                    {
                //                        InsertNewNodeInInterpolatedContourNodeList(el.NodeList[1], el.NodeList[0], ContourValue);
                //                    }
                //                    if (el.NodeList[1].Value >= ContourValue && el.NodeList[2].Value < ContourValue)
                //                    {
                //                        InsertNewNodeInInterpolatedContourNodeList(el.NodeList[1], el.NodeList[2], ContourValue);
                //                    }
                //                    if (el.NodeList[2].Value >= ContourValue && el.NodeList[1].Value < ContourValue)
                //                    {
                //                        InsertNewNodeInInterpolatedContourNodeList(el.NodeList[2], el.NodeList[1], ContourValue);
                //                    }
                //                    if (el.NodeList[2].Value >= ContourValue && el.NodeList[0].Value < ContourValue)
                //                    {
                //                        InsertNewNodeInInterpolatedContourNodeList(el.NodeList[2], el.NodeList[0], ContourValue);
                //                    }
                //                }
                //                else if (el.Type == 33)
                //                {
                //                    if (el.NodeList[0].Value >= ContourValue && el.NodeList[1].Value < ContourValue)
                //                    {
                //                        InsertNewNodeInInterpolatedContourNodeList(el.NodeList[0], el.NodeList[1], ContourValue);
                //                    }
                //                    if (el.NodeList[0].Value >= ContourValue && el.NodeList[3].Value < ContourValue)
                //                    {
                //                        InsertNewNodeInInterpolatedContourNodeList(el.NodeList[0], el.NodeList[3], ContourValue);
                //                    }
                //                    if (el.NodeList[1].Value >= ContourValue && el.NodeList[0].Value < ContourValue)
                //                    {
                //                        InsertNewNodeInInterpolatedContourNodeList(el.NodeList[1], el.NodeList[0], ContourValue);
                //                    }
                //                    if (el.NodeList[1].Value >= ContourValue && el.NodeList[2].Value < ContourValue)
                //                    {
                //                        InsertNewNodeInInterpolatedContourNodeList(el.NodeList[1], el.NodeList[2], ContourValue);
                //                    }
                //                    if (el.NodeList[2].Value >= ContourValue && el.NodeList[1].Value < ContourValue)
                //                    {
                //                        InsertNewNodeInInterpolatedContourNodeList(el.NodeList[2], el.NodeList[1], ContourValue);
                //                    }
                //                    if (el.NodeList[2].Value >= ContourValue && el.NodeList[3].Value < ContourValue)
                //                    {
                //                        InsertNewNodeInInterpolatedContourNodeList(el.NodeList[2], el.NodeList[3], ContourValue);
                //                    }
                //                    if (el.NodeList[3].Value >= ContourValue && el.NodeList[0].Value < ContourValue)
                //                    {
                //                        InsertNewNodeInInterpolatedContourNodeList(el.NodeList[3], el.NodeList[0], ContourValue);
                //                    }
                //                    if (el.NodeList[3].Value >= ContourValue && el.NodeList[2].Value < ContourValue)
                //                    {
                //                        InsertNewNodeInInterpolatedContourNodeList(el.NodeList[3], el.NodeList[2], ContourValue);
                //                    }
                //                }
                //                else
                //                {
                //                    UpdateTask(AppTaskID, "");
                //                    throw new Exception("Element type is not supported: Element type = [" + el.Type + "]");
                //                }
                //            }
                //            else
                //            {
                //                UpdateTask(AppTaskID, "");
                //                throw new Exception("Bottom does not exist outside the Dfsu3DSigma and Dfsu3DSigmaZ.");
                //            }
                //        }

                //        List<Node> UniqueNodeList = (from n in AllNodeList orderby n.ID select n).Distinct().ToList<Node>();

                //        ForwardVector = new Dictionary<String, Vector>();
                //        BackwardVector = new Dictionary<String, Vector>();

                //        // ------------------------- new code --------------------------
                //        //                     

                //        foreach (Element el in UniqueElementList)
                //        {
                //            if (el.Type == 21)
                //            {
                //                FillVectors21_32(el, UniqueElementList, ContourValue, AppTaskID, false, false);
                //            }
                //            else if (el.Type == 24)
                //            {
                //                UpdateTask(AppTaskID, "");
                //                throw new Exception("Element type is not supported: Element type = [" + el.Type + "]");
                //            }
                //            else if (el.Type == 25)
                //            {
                //                FillVectors25_33(el, UniqueElementList, ContourValue, AppTaskID, false, false);
                //            }
                //            else if (el.Type == 32)
                //            {
                //                FillVectors21_32(el, UniqueElementList, ContourValue, AppTaskID, true, false);
                //            }
                //            else if (el.Type == 33)
                //            {
                //                FillVectors25_33(el, UniqueElementList, ContourValue, AppTaskID, true, false);
                //            }
                //            else
                //            {
                //                UpdateTask(AppTaskID, "");
                //                throw new Exception("Element type is not supported: Element type = [" + el.Type + "]");
                //            }

                //        }

                //        //-------------- new code ------------------------


                //        bool MoreContourLine = true;
                //        while (MoreContourLine)
                //        {
                //            List<Node> FinalContourNodeList = new List<Node>();
                //            Vector LastVector = new Vector();
                //            LastVector = ForwardVector.First().Value;
                //            FinalContourNodeList.Add(LastVector.StartNode);
                //            FinalContourNodeList.Add(LastVector.EndNode);
                //            ForwardVector.Remove(LastVector.StartNode.ID.ToString() + "," + LastVector.EndNode.ID.ToString());
                //            BackwardVector.Remove(LastVector.EndNode.ID.ToString() + "," + LastVector.StartNode.ID.ToString());
                //            bool PolygonCompleted = false;
                //            while (!PolygonCompleted)
                //            {
                //                List<string> KeyStrList = (from k in ForwardVector.Keys
                //                                           where k.StartsWith(LastVector.EndNode.ID.ToString() + ",")
                //                                           && !k.EndsWith("," + LastVector.StartNode.ID.ToString())
                //                                           select k).ToList<string>();

                //                if (KeyStrList.Count != 1)
                //                {
                //                    KeyStrList = (from k in BackwardVector.Keys
                //                                  where k.StartsWith(LastVector.EndNode.ID.ToString() + ",")
                //                                  && !k.EndsWith("," + LastVector.StartNode.ID.ToString())
                //                                  select k).ToList<string>();

                //                    if (KeyStrList.Count != 1)
                //                    {
                //                        PolygonCompleted = true;
                //                        break;
                //                    }
                //                    else
                //                    {
                //                        LastVector = BackwardVector[KeyStrList[0]];
                //                        BackwardVector.Remove(LastVector.StartNode.ID.ToString() + "," + LastVector.EndNode.ID.ToString());
                //                        ForwardVector.Remove(LastVector.EndNode.ID.ToString() + "," + LastVector.StartNode.ID.ToString());
                //                    }
                //                }
                //                else
                //                {
                //                    LastVector = ForwardVector[KeyStrList[0]];
                //                    ForwardVector.Remove(LastVector.StartNode.ID.ToString() + "," + LastVector.EndNode.ID.ToString());
                //                    BackwardVector.Remove(LastVector.EndNode.ID.ToString() + "," + LastVector.StartNode.ID.ToString());
                //                }
                //                FinalContourNodeList.Add(LastVector.EndNode);
                //                if (FinalContourNodeList[FinalContourNodeList.Count - 1] == FinalContourNodeList[0])
                //                {
                //                    PolygonCompleted = true;
                //                }
                //            }

                //            if (CalculateAreaOfPolygon(FinalContourNodeList) < 0)
                //            {
                //                FinalContourNodeList.Reverse();
                //            }

                //            FinalContourNodeList.Add(FinalContourNodeList[0]);
                //            ContourPolygon contourPolygon = new ContourPolygon() { };
                //            contourPolygon.ContourNodeList = FinalContourNodeList;
                //            contourPolygon.ContourValue = ContourValue;
                //            ContourPolygonList.Add(contourPolygon);

                //            if (ForwardVector.Count == 0)
                //            {
                //                MoreContourLine = false;
                //            }
                //        }
                //        //sbKMLPollutionLimitsContour.AppendLine(@"<Folder>");
                //        //sbKMLPollutionLimitsContour.AppendLine(string.Format(@"<name>{0} Pollution Limits Contour</name>", ContourValue));
                //        //sbKMLPollutionLimitsContour.AppendLine(@"<visibility>0</visibility>");

                //        foreach (ContourPolygon contourPolygon in ContourPolygonList)
                //        {
                //            sbKMLPollutionLimitsContour.AppendLine(@"<Folder>");
                //            sbKMLPollutionLimitsContour.AppendLine(@"<visibility>0</visibility>");
                //            sbKMLPollutionLimitsContour.AppendLine(@"<Placemark>");
                //            sbKMLPollutionLimitsContour.AppendLine(@"<visibility>0</visibility>");
                //            if (contourPolygon.ContourValue >= 14 && contourPolygon.ContourValue < 88)
                //            {
                //                sbKMLPollutionLimitsContour.AppendLine(@"<styleUrl>#fc_14</styleUrl>");
                //            }
                //            else if (contourPolygon.ContourValue >= 88)
                //            {
                //                sbKMLPollutionLimitsContour.AppendLine(@"<styleUrl>#fc_88</styleUrl>");
                //            }
                //            else
                //            {
                //                sbKMLPollutionLimitsContour.AppendLine(@"<styleUrl>#fc_LT14</styleUrl>");
                //            }
                //            sbKMLPollutionLimitsContour.AppendLine(@"<Polygon>");
                //            sbKMLPollutionLimitsContour.AppendLine(@"<outerBoundaryIs>");
                //            sbKMLPollutionLimitsContour.AppendLine(@"<LinearRing>");
                //            sbKMLPollutionLimitsContour.AppendLine(@"<coordinates>");
                //            foreach (Node node in contourPolygon.ContourNodeList)
                //            {
                //                sbKMLPollutionLimitsContour.Append(string.Format(@"{0},{1},0 ", node.X, node.Y));
                //            }
                //            sbKMLPollutionLimitsContour.AppendLine(@"</coordinates>");
                //            sbKMLPollutionLimitsContour.AppendLine(@"</LinearRing>");
                //            sbKMLPollutionLimitsContour.AppendLine(@"</outerBoundaryIs>");
                //            sbKMLPollutionLimitsContour.AppendLine(@"</Polygon>");
                //            sbKMLPollutionLimitsContour.AppendLine(@"</Placemark>");
                //            sbKMLPollutionLimitsContour.AppendLine(@"</Folder>");
                //        }
                //        sbKMLPollutionLimitsContour.AppendLine(@"</Folder>");
                //        CountContour += 1;
                //    }
                //    sbKMLPollutionLimitsContour.AppendLine(@"</Folder>");
                //}
                #endregion Bottom of Layer
            }

            // Geneating Pollution Limits
            sbHTML.AppendLine(@"<Folder><name>" + TaskRunnerServiceRes.PollutionLimits + @"</name>");
            sbHTML.AppendLine(@"<visibility>0</visibility>");

            // Generating Polygons 
            sbHTML.AppendLine(@"<Folder><name>" + TaskRunnerServiceRes.Polygons + @"</name>");
            sbHTML.AppendLine(@"<visibility>0</visibility>");

            foreach (int Layer in LayerList)
            {
                if (Layer == 1)
                {
                    sbHTML.AppendLine(string.Format(@"<Folder><name>" + TaskRunnerServiceRes.TopOfLayer + @" [{0}] </name>", Layer));
                }
                else
                {
                    sbHTML.AppendLine(string.Format(@"<Folder><name>" + TaskRunnerServiceRes.TopOfLayer + @" [{0}] " + TaskRunnerServiceRes.BottomOfLayer + @" [{1}] </name>", Layer, Layer - 1));
                }
                sbHTML.AppendLine(@"<visibility>0</visibility>");

                foreach (List<ContourPolygon> contourPolygonList in ContourPolygonListList)
                {
                    foreach (ContourPolygon contourPolygon in contourPolygonList)
                    {
                        if (contourPolygon.Layer == Layer)
                        {
                            foreach (float contourValue in ContourValueList)
                            {
                                if (contourPolygon.ContourValue == contourValue)
                                {
                                    sbHTML.AppendLine(string.Format(@"<Folder><name>" + TaskRunnerServiceRes.ContourValue + @" [{0}]</name>", contourValue));

                                    sbHTML.AppendLine(@"<Folder>");
                                    sbHTML.AppendLine(@"<visibility>0</visibility>");
                                    sbHTML.AppendLine(@"<Placemark>");
                                    sbHTML.AppendLine(@"<visibility>0</visibility>");
                                    if (contourPolygon.ContourValue >= 14 && contourPolygon.ContourValue < 88)
                                    {
                                        sbHTML.AppendLine(@"<styleUrl>#fc_14</styleUrl>");
                                    }
                                    else if (contourPolygon.ContourValue >= 88)
                                    {
                                        sbHTML.AppendLine(@"<styleUrl>#fc_88</styleUrl>");
                                    }
                                    else
                                    {
                                        sbHTML.AppendLine(@"<styleUrl>#fc_LT14</styleUrl>");
                                    }
                                    sbHTML.AppendLine(@"<Polygon>");
                                    sbHTML.AppendLine(@"<outerBoundaryIs>");
                                    sbHTML.AppendLine(@"<LinearRing>");
                                    sbHTML.AppendLine(@"<coordinates>");
                                    foreach (Node node in contourPolygon.ContourNodeList)
                                    {
                                        sbHTML.Append(node.X.ToString().Replace(",", ".") + @"," + node.Y.ToString().Replace(",", ".") + "," + node.Z.ToString().Replace(",", ".") + " ");
                                    }
                                    sbHTML.AppendLine(@"</coordinates>");
                                    sbHTML.AppendLine(@"</LinearRing>");
                                    sbHTML.AppendLine(@"</outerBoundaryIs>");
                                    sbHTML.AppendLine(@"</Polygon>");
                                    sbHTML.AppendLine(@"</Placemark>");
                                    sbHTML.AppendLine(@"</Folder>");

                                    sbHTML.AppendLine(@"</Folder>");
                                }

                            }
                        }
                    }
                }
                sbHTML.AppendLine(@"</Folder>");
            }
            sbHTML.AppendLine(@"</Folder>");


            // Generating Depths 
            sbHTML.AppendLine(@"<Folder><name>" + TaskRunnerServiceRes.Depths + @"</name>");
            sbHTML.AppendLine(@"<visibility>0</visibility>");

            foreach (int Layer in LayerList)
            {
                if (Layer == 1)
                {
                    sbHTML.AppendLine(string.Format(@"<Folder><name>" + TaskRunnerServiceRes.TopOfLayer + @" [{0}] </name>", Layer));
                }
                else
                {
                    sbHTML.AppendLine(string.Format(@"<Folder><name>" + TaskRunnerServiceRes.TopOfLayer + @" [{0}] " + TaskRunnerServiceRes.BottomOfLayer + @" [{1}] </name>", Layer, Layer - 1));
                }
                sbHTML.AppendLine(@"<visibility>0</visibility>");

                foreach (List<ContourPolygon> contourPolygonList in ContourPolygonListList)
                {
                    foreach (ContourPolygon contourPolygon in contourPolygonList)
                    {
                        if (contourPolygon.Layer == Layer)
                        {
                            foreach (float contourValue in ContourValueList)
                            {
                                if (contourPolygon.ContourValue == contourValue)
                                {
                                    sbHTML.AppendLine(string.Format(@"<Folder><name>" + TaskRunnerServiceRes.ContourValue + @" [{0}]</name>", contourValue));

                                    sbHTML.AppendLine(@"<Folder>");
                                    sbHTML.AppendLine(@"<visibility>0</visibility>");
                                    sbHTML.AppendLine(@"<name>Depths</name>");
                                    foreach (Node node in contourPolygon.ContourNodeList)
                                    {
                                        sbHTML.AppendLine(@"<Placemark>");
                                        sbHTML.AppendLine(@"<visibility>0</visibility>");
                                        sbHTML.AppendLine(@"<name>" + node.Z.ToString("F1").Replace(",", ".") + "(m)</name>");
                                        if (contourPolygon.ContourValue >= 14 && contourPolygon.ContourValue < 88)
                                        {
                                            sbHTML.AppendLine(@"<styleUrl>#fc_14</styleUrl>");
                                        }
                                        else if (contourPolygon.ContourValue >= 88)
                                        {
                                            sbHTML.AppendLine(@"<styleUrl>#fc_88</styleUrl>");
                                        }
                                        else
                                        {
                                            sbHTML.AppendLine(@"<styleUrl>#fc_LT14</styleUrl>");
                                        }
                                        sbHTML.AppendLine(@"<Point>");
                                        sbHTML.AppendLine(@"<coordinates>");
                                        sbHTML.Append(node.X.ToString().Replace(",", ".") + @"," + node.Y.ToString().Replace(",", ".") + "," + node.Z.ToString().Replace(",", ".") + " ");
                                        sbHTML.AppendLine(@"</coordinates>");
                                        sbHTML.AppendLine(@"</Point>");
                                        sbHTML.AppendLine(@"</Placemark>");
                                    }
                                    sbHTML.AppendLine(@"</Folder>");

                                    sbHTML.AppendLine(@"</Folder>");
                                }
                            }
                        }
                    }
                }
                sbHTML.AppendLine(@"</Folder>");
            }
            sbHTML.AppendLine(@"</Folder>");

            sbHTML.AppendLine(@"</Folder>");

            string retStr2 = UpdateAppTaskPercentCompleted(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 100);

            return true;
        }
        private bool WriteKMLSourcePlacemark(StringBuilder sbHTML, PFSFile pfsFile, PFSSection pfsSectionSource, int SourceNumber, MikeSourceModel mikeSourceModel)
        {
            string NotUsed = "";
            string SourceName = GetSourceName(pfsSectionSource);
            if (string.IsNullOrWhiteSpace(SourceName))
            {
                NotUsed = TaskRunnerServiceRes.MIKESourceNameIsEmpty;
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageList("MIKESourceNameIsEmpty");
                return false;
            }

            int? SourceIncluded = GetSourceIncluded(pfsSectionSource);
            if (SourceIncluded == null)
            {
                return false;
            }

            float? SourceFlow = GetSourceFlow(pfsSectionSource);
            if (SourceFlow == null)
            {
                return false;
            }

            Coord SourceCoord = GetSourceCoord(pfsSectionSource);
            if (SourceCoord.Lat == 0.0f || SourceCoord.Lng == 0.0f)
            {
                return false;
            }

            List<Coord> coordList = new List<Coord>()
                    {
                        SourceCoord,
                    };

            PFSSection pfsSectionSourceTransport = pfsFile.GetSectionFromHandle("FemEngineHD/TRANSPORT_MODULE/SOURCES/SOURCE_" + SourceNumber.ToString() + "/COMPONENT_1");

            if (pfsSectionSourceTransport == null)
            {
                return false;
            }

            float? SourcePollution = GetSourcePollution(pfsSectionSourceTransport);
            if (SourcePollution == null)
            {
                return false;
            }

            int? SourcePollutionContinuous = GetSourcePollutionContinuous(pfsSectionSourceTransport);
            if (SourcePollutionContinuous == null)
            {
                return false;
            }

            PFSSection pfsSectionSourceTemperature = pfsFile.GetSectionFromHandle("FemEngineHD/HYDRODYNAMIC_MODULE/TEMPERATURE_SALINITY_MODULE/SOURCES/SOURCE_" + SourceNumber.ToString() + "/TEMPERATURE");

            if (pfsSectionSourceTemperature == null)
            {
                return false;
            }

            float? SourceTemperature = GetSourceTemperature(pfsSectionSourceTemperature);
            if (SourceTemperature == null)
            {
                return false;
            }

            PFSSection pfsSectionSourceSalinity = pfsFile.GetSectionFromHandle("FemEngineHD/HYDRODYNAMIC_MODULE/TEMPERATURE_SALINITY_MODULE/SOURCES/SOURCE_" + SourceNumber.ToString() + "/Salinity");

            if (pfsSectionSourceSalinity == null)
            {
                return false;
            }

            float? SourceSalinity = GetSourceSalinity(pfsSectionSourceSalinity);
            if (SourceSalinity == null)
            {
                return false;
            }


            sbHTML.Append("<Placemark>");

            if (SourceIncluded == 1)
            {
                if (mikeSourceModel.IsRiver)
                {
                    sbHTML.AppendLine(string.Format(@"<name>{0} (" + TaskRunnerServiceRes.UsedLowerCase + @")</name>", SourceName));
                    sbHTML.AppendLine(@"<styleUrl>#msn_blue-pushpin</styleUrl>");
                }
                else
                {
                    sbHTML.AppendLine(string.Format(@"<name>{0} (" + TaskRunnerServiceRes.UsedLowerCase + @")</name>", SourceName));
                    sbHTML.AppendLine(@"<styleUrl>#msn_grn-pushpin</styleUrl>");
                }
            }
            else
            {
                sbHTML.AppendLine(string.Format(@"<name>{0} (" + TaskRunnerServiceRes.NotUsedLowerCase + @")</name>", SourceName));
                sbHTML.AppendLine(@"<styleUrl>#msn_red-pushpin</styleUrl>");
            }
            sbHTML.AppendLine(@"<visibility>0</visibility>");
            sbHTML.AppendLine(@"<description><![CDATA[");
            sbHTML.AppendLine(string.Format(@"<h2>{0}</h2>", SourceName));
            sbHTML.AppendLine(@"<h3>" + TaskRunnerServiceRes.Effluent + @"</h3>");
            sbHTML.AppendLine(@"<ul>");
            sbHTML.AppendLine(@"<li><b>" + TaskRunnerServiceRes.Coordinates + @"</b>");
            sbHTML.AppendLine(string.Format(@"&nbsp;&nbsp;&nbsp; {0:F5} &nbsp; {1:F5}</li>", SourceCoord.Lat, SourceCoord.Lng).Replace(",", "."));

            List<MikeSourceStartEndModel> mikeSourceStartEndModelList = _MikeSourceStartEndService.GetMikeSourceStartEndModelListWithMikeSourceIDDB(mikeSourceModel.MikeSourceID);
            if ((bool)mikeSourceModel.IsContinuous)
            {
                if (mikeSourceStartEndModelList.Count > 0)
                {
                    sbHTML.AppendLine(@"<li><b>" + TaskRunnerServiceRes.IsContinuous + "</b></li>");
                    sbHTML.AppendLine(string.Format(@"<li><b>" + TaskRunnerServiceRes.Flow + @":</b> {0:F6} (m3/s)  {1:F0} (m3/" + TaskRunnerServiceRes.DayLowerCase + @")</li>", mikeSourceStartEndModelList[0].SourceFlowStart_m3_day / 24 / 3600, mikeSourceStartEndModelList[0].SourceFlowStart_m3_day).ToString().Replace(",", "."));
                    sbHTML.AppendLine(string.Format(@"<li><b>" + TaskRunnerServiceRes.FCMPNPer100ML + @":</b> {0:F0}</li>", mikeSourceStartEndModelList[0].SourcePollutionStart_MPN_100ml).ToString().Replace(",", "."));
                    sbHTML.AppendLine(string.Format(@"<li><b>" + TaskRunnerServiceRes.Temperature + @":</b> {0:F0} " + TaskRunnerServiceRes.Celcius + @"</li>", mikeSourceStartEndModelList[0].SourceTemperatureStart_C).ToString().Replace(",", "."));
                    sbHTML.AppendLine(string.Format(@"<li><b>" + TaskRunnerServiceRes.Salinity + @":</b> {0:F0} " + TaskRunnerServiceRes.PSU + @"</li>", mikeSourceStartEndModelList[0].SourceSalinityStart_PSU).ToString().Replace(",", "."));
                }
            }
            else
            {
                sbHTML.AppendLine(@"<li><b>" + TaskRunnerServiceRes.IsNotContinuous + @"</b></li>");

                mikeSourceStartEndModelList = _MikeSourceStartEndService.GetMikeSourceStartEndModelListWithMikeSourceIDDB(mikeSourceModel.MikeSourceID);
                int CountMikeSourceStartEnd = 0;
                foreach (MikeSourceStartEndModel mssem in mikeSourceStartEndModelList)
                {
                    CountMikeSourceStartEnd += 1;
                    sbHTML.AppendLine(@"<ul>");
                    sbHTML.AppendLine("<b>" + TaskRunnerServiceRes.Spill + @" " + CountMikeSourceStartEnd + "</b>");
                    sbHTML.AppendLine(string.Format(@"<li><b>" + TaskRunnerServiceRes.SpillStartTime + @":</b> {0:yyyy/MM/dd HH:mm:ss tt}</li>", mssem.StartDateAndTime_Local));
                    sbHTML.AppendLine(string.Format(@"<li><b>" + TaskRunnerServiceRes.SpillEndTime + @":</b> {0:yyyy/MM/dd HH:mm:ss tt}</li>", mssem.EndDateAndTime_Local));
                    sbHTML.AppendLine(string.Format(@"<li><b>" + TaskRunnerServiceRes.FlowStart + @":</b> {0:F6} (m3/s)  {1:F0} (m3/" + TaskRunnerServiceRes.DayLowerCase + @")</li>", mssem.SourceFlowStart_m3_day / 24 / 3600, mssem.SourceFlowStart_m3_day).ToString().Replace(",", "."));
                    sbHTML.AppendLine(string.Format(@"<li><b>" + TaskRunnerServiceRes.FlowEnd + @":</b> {0:F6} (m3/s)  {1:F0} (m3/" + TaskRunnerServiceRes.DayLowerCase + @")</li>", mssem.SourceFlowEnd_m3_day / 24 / 3600, mssem.SourceFlowEnd_m3_day).ToString().Replace(",", "."));
                    sbHTML.AppendLine(string.Format(@"<li><b>" + TaskRunnerServiceRes.FCMPNPer100MLStart + @":</b> {0:F0}</li>", mssem.SourcePollutionStart_MPN_100ml).ToString().Replace(",", "."));
                    sbHTML.AppendLine(string.Format(@"<li><b>" + TaskRunnerServiceRes.FCMPNPer100MLEnd + @":</b> {0:F0}</li>", mssem.SourcePollutionEnd_MPN_100ml).ToString().Replace(",", "."));
                    sbHTML.AppendLine(string.Format(@"<li><b>" + TaskRunnerServiceRes.TemperatureStart + @":</b> {0:F0} " + TaskRunnerServiceRes.Celcius + @"</li>", mssem.SourceTemperatureStart_C).ToString().Replace(",", "."));
                    sbHTML.AppendLine(string.Format(@"<li><b>" + TaskRunnerServiceRes.TemperatureEnd + @":</b> {0:F0} " + TaskRunnerServiceRes.Celcius + @"</li>", mssem.SourceTemperatureEnd_C).ToString().Replace(",", "."));
                    sbHTML.AppendLine(string.Format(@"<li><b>" + TaskRunnerServiceRes.SalinityStart + @":</b> {0:F0} " + TaskRunnerServiceRes.PSU + @"</li>", mssem.SourceSalinityStart_PSU).ToString().Replace(",", "."));
                    sbHTML.AppendLine(string.Format(@"<li><b>" + TaskRunnerServiceRes.SalinityEnd + @":</b> {0:F0} " + TaskRunnerServiceRes.PSU + @"</li>", mssem.SourceSalinityEnd_PSU).ToString().Replace(",", "."));
                    sbHTML.AppendLine(@"</ul>");
                }
            }
            sbHTML.AppendLine(@"</ul>");
            sbHTML.AppendLine(@"<iframe src=""about:"" width=""500"" height=""1"" />");
            sbHTML.AppendLine(@"]]></description>");

            sbHTML.AppendLine(@"<Point>");
            sbHTML.AppendLine(@"<coordinates>");
            sbHTML.AppendLine(SourceCoord.Lng.ToString().Replace(",", ".") + @"," + SourceCoord.Lat.ToString().Replace(",", ".") + ",0 ");
            sbHTML.AppendLine(@"</coordinates>");
            sbHTML.AppendLine(@"</Point>");
            sbHTML.AppendLine(@"</Placemark>");

            return true;
        }
        private bool WriteKMLStudyAreaLine(StringBuilder sbHTML, List<ElementLayer> elementLayerList, List<Node> nodeList)
        {
            sbHTML.AppendLine(@"  <Style id=""StudyArea"">");
            sbHTML.AppendLine(@"  <LineStyle>");
            sbHTML.AppendLine(@"  <color>ffffff00</color>");
            sbHTML.AppendLine(@"  <width>2</width>");
            sbHTML.AppendLine(@"  </LineStyle>");
            sbHTML.AppendLine(@"  </Style>");

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


            //DrawKMLUniqueNodes(UniqueNodeList, 0, sbStyleStudyAreaLine, sbKMLStudyAreaLine);
            //DrawKMLInterpolatedContourNodes(InterpolatedContourNodeList, sbStyleStudyAreaLine, sbKMLStudyAreaLine);
            //DrawUniqueElements(UniqueElementList, sbStyleStudyAreaLine, sbKMLStudyAreaLine);

            //UpdateTask(AppTaskID, "60 %");

            bool MoreStudyAreaLine = true;
            int PolygonCount = 0;
            while (MoreStudyAreaLine)
            {
                PolygonCount += 1;

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

                double Area = _MapInfoService.CalculateAreaOfPolygon(FinalContourNodeList);
                if (Area < 0)
                {
                    FinalContourNodeList.Reverse();
                    Area = _MapInfoService.CalculateAreaOfPolygon(FinalContourNodeList);
                }

                FinalContourNodeList.Add(FinalContourNodeList[0]);

                ContourPolygon contourLine = new ContourPolygon() { };
                contourLine.ContourNodeList = FinalContourNodeList;
                contourLine.ContourValue = 0;
                ContourPolygonList.Add(contourLine);

                if (ForwardVector.Count == 0)
                {
                    MoreStudyAreaLine = false;
                }

            }

            sbHTML.AppendLine(@"  <Folder>");
            sbHTML.AppendLine(@"    <name>" + TaskRunnerServiceRes.MIKEStudyArea + @"</name>");
            sbHTML.AppendLine(@"    <visibility>0</visibility>");
            foreach (ContourPolygon contourPolygon in ContourPolygonList)
            {
                sbHTML.AppendLine(@"    <Folder>");
                sbHTML.AppendLine(@"      <visibility>0</visibility>");
                sbHTML.AppendLine(@"      <Placemark>");
                sbHTML.AppendLine(@"        <visibility>0</visibility>");
                sbHTML.AppendLine(@"        <styleUrl>#StudyArea</styleUrl>");
                sbHTML.AppendLine(@"        <LineString>");
                sbHTML.AppendLine(@"        <coordinates>");
                foreach (Node node in contourPolygon.ContourNodeList)
                {
                    sbHTML.Append("        " + node.X.ToString().Replace(",", ".") + @"," + node.Y.ToString().Replace(",", ".") + ",0 ");
                }
                sbHTML.AppendLine(@"        </coordinates>");
                sbHTML.AppendLine(@"        </LineString>");
                sbHTML.AppendLine(@"      </Placemark>");
                sbHTML.AppendLine(@"    </Folder>");
            }
            sbHTML.AppendLine(@"  </Folder>");

            return true;
        }
        private bool WriteKMLStyleModelInput(StringBuilder sbHTML)
        {
            sbHTML.AppendLine(@"<Style id=""sn_grn-pushpin"">");
            sbHTML.AppendLine(@"<IconStyle>");
            sbHTML.AppendLine(@"<scale>1.1</scale>");
            sbHTML.AppendLine(@"<Icon>");
            sbHTML.AppendLine(@"<href>http://maps.google.com/mapfiles/kml/pushpin/grn-pushpin.png</href>");
            sbHTML.AppendLine(@"</Icon>");
            sbHTML.AppendLine(@"<hotSpot x=""20"" y=""2"" xunits=""pixels"" yunits=""pixels""/>");
            sbHTML.AppendLine(@"</IconStyle>");
            sbHTML.AppendLine(@"<ListStyle>");
            sbHTML.AppendLine(@"</ListStyle>");
            sbHTML.AppendLine(@"</Style>");

            sbHTML.AppendLine(@"<Style id=""sh_grn-pushpin"">");
            sbHTML.AppendLine(@"<IconStyle>");
            sbHTML.AppendLine(@"<scale>1.3</scale>");
            sbHTML.AppendLine(@"<Icon>");
            sbHTML.AppendLine(@"<href>http://maps.google.com/mapfiles/kml/pushpin/grn-pushpin.png</href>");
            sbHTML.AppendLine(@"</Icon>");
            sbHTML.AppendLine(@"<hotSpot x=""20"" y=""2"" xunits=""pixels"" yunits=""pixels""/>");
            sbHTML.AppendLine(@"</IconStyle>");
            sbHTML.AppendLine(@"<ListStyle>");
            sbHTML.AppendLine(@"</ListStyle>");
            sbHTML.AppendLine(@"</Style>");

            sbHTML.AppendLine(@"<StyleMap id=""msn_grn-pushpin"">");
            sbHTML.AppendLine(@"<Pair>");
            sbHTML.AppendLine(@"<key>normal</key>");
            sbHTML.AppendLine(@"<styleUrl>#sn_grn-pushpin</styleUrl>");
            sbHTML.AppendLine(@"</Pair>");
            sbHTML.AppendLine(@"<Pair>");
            sbHTML.AppendLine(@"<key>highlight</key>");
            sbHTML.AppendLine(@"<styleUrl>#sh_grn-pushpin</styleUrl>");
            sbHTML.AppendLine(@"</Pair>");
            sbHTML.AppendLine(@"</StyleMap>");

            sbHTML.AppendLine(@"<Style id=""sn_red-pushpin"">");
            sbHTML.AppendLine(@"<IconStyle>");
            sbHTML.AppendLine(@"<scale>1.1</scale>");
            sbHTML.AppendLine(@"<Icon>");
            sbHTML.AppendLine(@"<href>http://maps.google.com/mapfiles/kml/pushpin/red-pushpin.png</href>");
            sbHTML.AppendLine(@"</Icon>");
            sbHTML.AppendLine(@"<hotSpot x=""20"" y=""2"" xunits=""pixels"" yunits=""pixels""/>");
            sbHTML.AppendLine(@"</IconStyle>");
            sbHTML.AppendLine(@"<ListStyle>");
            sbHTML.AppendLine(@"</ListStyle>");
            sbHTML.AppendLine(@"</Style>");

            sbHTML.AppendLine(@"<Style id=""sh_red-pushpin"">");
            sbHTML.AppendLine(@"<IconStyle>");
            sbHTML.AppendLine(@"<scale>1.3</scale>");
            sbHTML.AppendLine(@"<Icon>");
            sbHTML.AppendLine(@"<href>http://maps.google.com/mapfiles/kml/pushpin/red-pushpin.png</href>");
            sbHTML.AppendLine(@"</Icon>");
            sbHTML.AppendLine(@"<hotSpot x=""20"" y=""2"" xunits=""pixels"" yunits=""pixels""/>");
            sbHTML.AppendLine(@"</IconStyle>");
            sbHTML.AppendLine(@"<ListStyle>");
            sbHTML.AppendLine(@"</ListStyle>");
            sbHTML.AppendLine(@"</Style>");

            sbHTML.AppendLine(@"<StyleMap id=""msn_red-pushpin"">");
            sbHTML.AppendLine(@"<Pair>");
            sbHTML.AppendLine(@"<key>normal</key>");
            sbHTML.AppendLine(@"<styleUrl>#sn_red-pushpin</styleUrl>");
            sbHTML.AppendLine(@"</Pair>");
            sbHTML.AppendLine(@"<Pair>");
            sbHTML.AppendLine(@"<key>highlight</key>");
            sbHTML.AppendLine(@"<styleUrl>#sh_red-pushpin</styleUrl>");
            sbHTML.AppendLine(@"</Pair>");
            sbHTML.AppendLine(@"</StyleMap>");

            sbHTML.AppendLine(@"<Style id=""sn_blue-pushpin"">");
            sbHTML.AppendLine(@"<IconStyle>");
            sbHTML.AppendLine(@"<scale>1.1</scale>");
            sbHTML.AppendLine(@"<Icon>");
            sbHTML.AppendLine(@"<href>http://maps.google.com/mapfiles/kml/pushpin/blue-pushpin.png</href>");
            sbHTML.AppendLine(@"</Icon>");
            sbHTML.AppendLine(@"<hotSpot x=""20"" y=""2"" xunits=""pixels"" yunits=""pixels""/>");
            sbHTML.AppendLine(@"</IconStyle>");
            sbHTML.AppendLine(@"<ListStyle>");
            sbHTML.AppendLine(@"</ListStyle>");
            sbHTML.AppendLine(@"</Style>");

            sbHTML.AppendLine(@"<Style id=""sh_blue-pushpin"">");
            sbHTML.AppendLine(@"<IconStyle>");
            sbHTML.AppendLine(@"<scale>1.3</scale>");
            sbHTML.AppendLine(@"<Icon>");
            sbHTML.AppendLine(@"<href>http://maps.google.com/mapfiles/kml/pushpin/blue-pushpin.png</href>");
            sbHTML.AppendLine(@"</Icon>");
            sbHTML.AppendLine(@"<hotSpot x=""20"" y=""2"" xunits=""pixels"" yunits=""pixels""/>");
            sbHTML.AppendLine(@"</IconStyle>");
            sbHTML.AppendLine(@"<ListStyle>");
            sbHTML.AppendLine(@"</ListStyle>");
            sbHTML.AppendLine(@"</Style>");

            sbHTML.AppendLine(@"<StyleMap id=""msn_blue-pushpin"">");
            sbHTML.AppendLine(@"<Pair>");
            sbHTML.AppendLine(@"<key>normal</key>");
            sbHTML.AppendLine(@"<styleUrl>#sn_blue-pushpin</styleUrl>");
            sbHTML.AppendLine(@"</Pair>");
            sbHTML.AppendLine(@"<Pair>");
            sbHTML.AppendLine(@"<key>highlight</key>");
            sbHTML.AppendLine(@"<styleUrl>#sh_blue-pushpin</styleUrl>");
            sbHTML.AppendLine(@"</Pair>");
            sbHTML.AppendLine(@"</StyleMap>");

            return true;
        }
        private bool WriteKMLTop(StringBuilder sbHTML)
        {
            sbHTML.AppendLine(@"<?xml version=""1.0"" encoding=""UTF-8""?>");
            sbHTML.AppendLine(@"<kml xmlns=""http://www.opengis.net/kml/2.2"" xmlns:gx=""http://www.google.com/kml/ext/2.2"" xmlns:kml=""http://www.opengis.net/kml/2.2"" xmlns:atom=""http://www.w3.org/2005/Atom"">");
            sbHTML.AppendLine(@"<Document>");

            return true;
        }

    }
}
