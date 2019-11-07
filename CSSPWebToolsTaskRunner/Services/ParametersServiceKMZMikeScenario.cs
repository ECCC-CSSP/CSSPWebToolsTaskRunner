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
using System.Threading;
using System.Globalization;
using DHI.Generic.MikeZero.DFS.dfsu;
using DHI.PFS;
using DHI.Generic.MikeZero.DFS;
using DHI.Generic.MikeZero;
using System.Drawing;
using System.Xml;

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
                case "MikeScenarioCurrentsAnimationKMZ":
                    {
                        if (!GenerateMikeScenarioCurrentsAnimationKMZ())
                        {
                            return false;
                        }
                    }
                    break;
                case "MikeScenarioEstimatedDroguePathsAnimationKMZ":
                    {
                        if (!GenerateMikeScenarioEstimatedDroguePathsAnimationKMZ())
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
                NotUsed = string.Format(TaskRunnerServiceRes.ErrorOccuredWhileCreating_Document_, "KMZ", fi.FullName);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("ErrorOccuredWhileCreating_Document_", "KMZ", fi.FullName);
                return false;
            }

            return true;
        }
        public bool GenerateMikeScenarioCurrentsAnimationKMZ()
        {
            string NotUsed = "";
            bool ErrorInDoc = false;

            List<string> ProvinceTextEN = new List<string>()
            {
                "New Brunswick",
                "Newfoundland and Labrador",
                "Nova Scotia",
                "Prince Edward Island",
                "Québec",
                "British Columbia",
                "Maine",
                "Washington"
            };

            List<string> ProvinceTextFR = new List<string>()
            {
                "Nouveau-Brunswick",
                "Terre-Neuve-et-Labrador",
                "Nouvelle-Écosse",
                "Île-du-Prince-Édouard",
                "Québec",
                "Colombie-Britannique",
                "Maine",
                "Washington"
            };

            string GoogleEarthPath = "";
            List<string> ParamValueList = Parameters.Split("|||".ToCharArray(), StringSplitOptions.RemoveEmptyEntries).ToList();

            GoogleEarthPath = GetParameters("GoogleEarthPath", ParamValueList);
            if (string.IsNullOrWhiteSpace(GoogleEarthPath))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind__, TaskRunnerServiceRes.Parameter, TaskRunnerServiceRes.GoogleEarthPath);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotFind__", TaskRunnerServiceRes.Parameter, TaskRunnerServiceRes.GoogleEarthPath);
                return false;
            }

            GoogleEarthPath = GoogleEarthPath.Replace("!!!!!", "<");
            GoogleEarthPath = GoogleEarthPath.Replace("@@@@@", ">");
            GoogleEarthPath = GoogleEarthPath.Replace("%%%%%", ",");

            TVItemModel tvItemModelMikeScenario = _TVItemService.GetTVItemModelWithTVItemIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID);

            List<TVItemModel> tvItemModelParentList = _TVItemService.GetParentsTVItemModelList(tvItemModelMikeScenario.TVPath);

            TVItemModel tvItemModelProvince = null;
            foreach (TVItemModel tvItemModel in tvItemModelParentList)
            {
                if (tvItemModel.TVType == TVTypeEnum.Province)
                {
                    tvItemModelProvince = tvItemModel;
                }
            }

            if (tvItemModelProvince == null)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_, TaskRunnerServiceRes.Province + "," + TaskRunnerServiceRes.TVItem);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotFind_", TaskRunnerServiceRes.Province + "," + TaskRunnerServiceRes.TVItem);
                return false;
            }

            string ProvInit = "";
            for (int i = 0; i < ProvinceTextEN.Count; i++)
            {
                if (ProvinceTextEN[i] == tvItemModelProvince.TVText || ProvinceTextFR[i] == tvItemModelProvince.TVText)
                {
                    switch (i)
                    {
                        case 0:
                            ProvInit = "NB";
                            break;
                        case 1:
                            ProvInit = "NL";
                            break;
                        case 2:
                            ProvInit = "NS";
                            break;
                        case 3:
                            ProvInit = "PE";
                            break;
                        case 4:
                            ProvInit = "QC";
                            break;
                        case 5:
                            ProvInit = "BC";
                            break;
                        case 6:
                            ProvInit = "ME";
                            break;
                        case 7:
                            ProvInit = "WA";
                            break;
                        default:
                            break;
                    }
                }
            }

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
                    NotUsed = string.Format(TaskRunnerServiceRes.Error_WhileCreating_Document_, "GetTransportDfsuFile", "KMZ", fi.FullName);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("Error_WhileCreating_Document_", "GetTransportDfsuFile", "KMZ", fi.FullName);
                }
                return false;
            }
            if (!FillRequiredList(dfsuFile, elementList, elementLayerList, nodeList, topNodeLayerList, bottomNodeLayerList))
            {
                if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.Error_WhileCreating_Document_, "FillRequiredList", fi.FullName);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("Error_WhileCreatingKMZDocument_", "FillRequiredList", "KMZ", fi.FullName);
                }
                return false;
            }

            List<ElementLayer> SelectedElementLayerList = ParseKMLPath(GoogleEarthPath, elementLayerList);
            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
                return false;

            if (!WriteKMLTop(sb))
            {
                ErrorInDoc = true;
            }
            sb.AppendLine(string.Format(@"<name>{0}</name>", TaskRunnerServiceRes.MIKECurrentsAnimation + "_" + GetScenarioName()));

            if (!DrawKMLCurrentsStyle(sb))
            {
                ErrorInDoc = true;
            }

            if (!WriteKMLCurrentsAnimation(sb, dfsuFile, elementLayerList, topNodeLayerList, bottomNodeLayerList, SelectedElementLayerList, ProvInit))
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
                NotUsed = string.Format(TaskRunnerServiceRes.ErrorOccuredWhileCreating_Document_, "KMZ", fi.FullName);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("ErrorOccuredWhileCreating_Document_", "KMZ", fi.FullName);
                return false;
            }
            return true;
        }
        public bool GenerateMikeScenarioEstimatedDroguePathsAnimationKMZ()
        {
            string NotUsed = "";
            bool ErrorInDoc = false;

            List<string> ProvinceTextEN = new List<string>()
            {
                "New Brunswick",
                "Newfoundland and Labrador",
                "Nova Scotia",
                "Prince Edward Island",
                "Québec",
                "British Columbia",
                "Maine",
                "Washington"
            };

            List<string> ProvinceTextFR = new List<string>()
            {
                "Nouveau-Brunswick",
                "Terre-Neuve-et-Labrador",
                "Nouvelle-Écosse",
                "Île-du-Prince-Édouard",
                "Québec",
                "Colombie-Britannique",
                "Maine",
                "Washington"
            };

            int DoFirstXDroguePoints = 0;
            List<int> DelaysList = new List<int>();
            List<int> LayersList = new List<int>();
            string GoogleEarthPath = "";

            List<string> ParamValueList = Parameters.Split("|||".ToCharArray(), StringSplitOptions.RemoveEmptyEntries).ToList();

            DoFirstXDroguePoints = int.Parse(GetParameters("DoFirstXDroguePoints", ParamValueList));

            List<string> DelaysTextList = GetParametersArray("Delays", ParamValueList);
            if (DelaysTextList.Count > 0)
            {
                foreach (string DelaysText in DelaysTextList)
                {
                    DelaysList.Add(int.Parse(DelaysText));
                }
            }

            List<string> LayersTextList = GetParametersArray("Layers", ParamValueList);
            if (LayersTextList.Count > 0)
            {
                foreach (string LayersText in LayersTextList)
                {
                    LayersList.Add(int.Parse(LayersText));
                }
            }

            GoogleEarthPath = GetParameters("GoogleEarthPath", ParamValueList);
            if (string.IsNullOrWhiteSpace(GoogleEarthPath))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind__, TaskRunnerServiceRes.Parameter, TaskRunnerServiceRes.GoogleEarthPath);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("CouldNotFind__", TaskRunnerServiceRes.Parameter, TaskRunnerServiceRes.GoogleEarthPath);
                return false;
            }

            GoogleEarthPath = GoogleEarthPath.Replace("!!!!!", "<");
            GoogleEarthPath = GoogleEarthPath.Replace("@@@@@", ">");
            GoogleEarthPath = GoogleEarthPath.Replace("%%%%%", ",");

            TVItemModel tvItemModelMikeScenario = _TVItemService.GetTVItemModelWithTVItemIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID);

            List<TVItemModel> tvItemModelParentList = _TVItemService.GetParentsTVItemModelList(tvItemModelMikeScenario.TVPath);

            TVItemModel tvItemModelProvince = null;
            foreach (TVItemModel tvItemModel in tvItemModelParentList)
            {
                if (tvItemModel.TVType == TVTypeEnum.Province)
                {
                    tvItemModelProvince = tvItemModel;
                }
            }

            if (tvItemModelProvince == null)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_, TaskRunnerServiceRes.Province + "," + TaskRunnerServiceRes.TVItem);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotFind_", TaskRunnerServiceRes.Province + "," + TaskRunnerServiceRes.TVItem);
                return false;
            }

            string ProvInit = "";
            for (int i = 0; i < ProvinceTextEN.Count; i++)
            {
                if (ProvinceTextEN[i] == tvItemModelProvince.TVText || ProvinceTextFR[i] == tvItemModelProvince.TVText)
                {
                    switch (i)
                    {
                        case 0:
                            ProvInit = "NB";
                            break;
                        case 1:
                            ProvInit = "NL";
                            break;
                        case 2:
                            ProvInit = "NS";
                            break;
                        case 3:
                            ProvInit = "PE";
                            break;
                        case 4:
                            ProvInit = "QC";
                            break;
                        case 5:
                            ProvInit = "BC";
                            break;
                        case 6:
                            ProvInit = "ME";
                            break;
                        case 7:
                            ProvInit = "WA";
                            break;
                        default:
                            break;
                    }
                }
            }

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
                    NotUsed = string.Format(TaskRunnerServiceRes.Error_WhileCreating_Document_, "GetTransportDfsuFile", "KMZ", fi.FullName);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("Error_WhileCreating_Document_", "GetTransportDfsuFile", "KMZ", fi.FullName);
                }
                return false;
            }
            if (!FillRequiredList(dfsuFile, elementList, elementLayerList, nodeList, topNodeLayerList, bottomNodeLayerList))
            {
                if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.Error_WhileCreating_Document_, "FillRequiredList", fi.FullName);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("Error_WhileCreatingKMZDocument_", "FillRequiredList", "KMZ", fi.FullName);
                }
                return false;
            }

            List<ElementLayer> SelectedElementLayerList = ParseKMLPath(GoogleEarthPath, elementLayerList);
            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
                return false;

            List<Coord> StartCoordList = ParseKMLPathCoord(GoogleEarthPath, elementLayerList);
            if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count > 0)
                return false;

            if (!WriteKMLTop(sb))
            {
                ErrorInDoc = true;
            }
            sb.AppendLine(string.Format(@"<name>{0}</name>", TaskRunnerServiceRes.EstimatedDroguePathsAnim + "_" + GetScenarioName()));

            if (!DrawKMLEstimatedDroguePathsStyle(sb))
            {
                ErrorInDoc = true;
            }

            if (!WriteKMLEstimatedDroguePathsAnimation(sb, dfsuFile, elementLayerList, topNodeLayerList, bottomNodeLayerList, SelectedElementLayerList, StartCoordList, DoFirstXDroguePoints, DelaysList, LayersList, ProvInit))
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
                NotUsed = string.Format(TaskRunnerServiceRes.ErrorOccuredWhileCreating_Document_, "KMZ", fi.FullName);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("ErrorOccuredWhileCreating_Document_", "KMZ", fi.FullName);
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
                    NotUsed = string.Format(TaskRunnerServiceRes.Error_WhileCreating_Document_, "GettingHydrodynamicDfsuFile", "KMZ", fi.FullName);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("Error_WhileCreating_Document_", "GettingHydrodynamicDfsuFile", "KMZ", fi.FullName);
                }
                return false;
            }
            if (!FillRequiredList(dfsuFile, elementList, elementLayerList, nodeList, topNodeLayerList, bottomNodeLayerList))
            {
                if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.Error_WhileCreating_Document_, "FillRequiredList", "KMZ", fi.FullName);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("Error_WhileCreating_Document_", "FillRequiredList", "KMZ", fi.FullName);
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
                NotUsed = string.Format(TaskRunnerServiceRes.ErrorOccuredWhileCreating_Document_, fi.FullName);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("ErrorOccuredWhileCreating_Document_", "KMZ", fi.FullName);
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
                    NotUsed = string.Format(TaskRunnerServiceRes.Error_WhileCreating_Document_, "GettingHydrodynamicDfsuFile", "KMZ", fi.FullName);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("Error_WhileCreating_Document_", "GettingHydrodynamicDfsuFile", "KMZ", fi.FullName);
                }
                return false;
            }
            if (!FillRequiredList(dfsuFile, elementList, elementLayerList, nodeList, topNodeLayerList, bottomNodeLayerList))
            {
                if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.Error_WhileCreating_Document_, "FillRequiredList", "KMZ", fi.FullName);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("Error_WhileCreating_Document_", "FillRequiredList", "KMZ", fi.FullName);
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
                NotUsed = string.Format(TaskRunnerServiceRes.ErrorOccuredWhileCreating_Document_, "KMZ", fi.FullName);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("ErrorOccuredWhileCreating_Document_", "KMZ", fi.FullName);
                return false;
            }

            return true;
        }
        public bool GenerateMikeScenarioPollutionAnimationKMZ()
        {
            string NotUsed = "";
            bool ErrorInDoc = false;
            List<string> ProvinceTextEN = new List<string>()
            {
                "New Brunswick",
                "Newfoundland and Labrador",
                "Nova Scotia",
                "Prince Edward Island",
                "Québec",
                "British Columbia",
                "Maine",
                "Washington"
            };

            List<string> ProvinceTextFR = new List<string>()
            {
                "Nouveau-Brunswick",
                "Terre-Neuve-et-Labrador",
                "Nouvelle-Écosse",
                "Île-du-Prince-Édouard",
                "Québec",
                "Colombie-Britannique",
                "Maine",
                "Washington"
            };

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

            TVItemModel tvItemModelMikeScenario = _TVItemService.GetTVItemModelWithTVItemIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.TVItemID);

            List<TVItemModel> tvItemModelParentList = _TVItemService.GetParentsTVItemModelList(tvItemModelMikeScenario.TVPath);

            TVItemModel tvItemModelProvince = null;
            foreach (TVItemModel tvItemModel in tvItemModelParentList)
            {
                if (tvItemModel.TVType == TVTypeEnum.Province)
                {
                    tvItemModelProvince = tvItemModel;
                }
            }

            if (tvItemModelProvince == null)
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_, TaskRunnerServiceRes.Province + "," + TaskRunnerServiceRes.TVItem);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotFind_", TaskRunnerServiceRes.Province + "," + TaskRunnerServiceRes.TVItem);
                return false;
            }

            string ProvInit = "";
            for (int i = 0; i < ProvinceTextEN.Count; i++)
            {
                if (ProvinceTextEN[i] == tvItemModelProvince.TVText || ProvinceTextFR[i] == tvItemModelProvince.TVText)
                {
                    switch (i)
                    {
                        case 0:
                            ProvInit = "NB";
                            break;
                        case 1:
                            ProvInit = "NL";
                            break;
                        case 2:
                            ProvInit = "NS";
                            break;
                        case 3:
                            ProvInit = "PE";
                            break;
                        case 4:
                            ProvInit = "QC";
                            break;
                        case 5:
                            ProvInit = "BC";
                            break;
                        case 6:
                            ProvInit = "ME";
                            break;
                        case 7:
                            ProvInit = "WA";
                            break;
                        default:
                            break;
                    }
                }
            }

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
                    NotUsed = string.Format(TaskRunnerServiceRes.Error_WhileCreating_Document_, "GetTransportDfsuFile", "KMZ", fi.FullName);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("Error_WhileCreating_Document_", "GetTransportDfsuFile", "KMZ", fi.FullName);
                }
                return false;
            }
            if (!FillRequiredList(dfsuFile, elementList, elementLayerList, nodeList, topNodeLayerList, bottomNodeLayerList))
            {
                if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.Error_WhileCreating_Document_, "FillRequiredList", "KMZ", fi.FullName);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("Error_WhileCreating_Document_", "FillRequiredList", "KMZ", fi.FullName);
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
            if (!WriteKMLFecalColiformContourLine(sb, dfsuFile, elementLayerList, topNodeLayerList, bottomNodeLayerList, ContourValueList, ProvInit))
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
                NotUsed = string.Format(TaskRunnerServiceRes.ErrorOccuredWhileCreating_Document_, "KMZ", fi.FullName);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("ErrorOccuredWhileCreating_Document_", "KMZ", fi.FullName);
                return false;
            }
            return true;
        }
        public bool GenerateMikeScenarioPollutionLimitKMZ()
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
                    NotUsed = string.Format(TaskRunnerServiceRes.Error_WhileCreating_Document_, "GetTransportDfsuFile", "KMZ", fi.FullName);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("Error_WhileCreating_Document_", "GetTransportDfsuFile", "KMZ", fi.FullName);
                }
                return false;
            }
            if (!FillRequiredList(dfsuFile, elementList, elementLayerList, nodeList, topNodeLayerList, bottomNodeLayerList))
            {
                if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.Error_WhileCreating_Document_, "FillRequiredList", "KMZ", fi.FullName);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("Error_WhileCreatingKMZDocument_", "FillRequiredList", "KMZ", fi.FullName);
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
                NotUsed = string.Format(TaskRunnerServiceRes.ErrorOccuredWhileCreating_Document_, "KMZ", fi.FullName);
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat2List("ErrorOccuredWhileCreating_Document_", "KMZ", fi.FullName);
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
        private void DrawKMLContourPolygon(List<ContourPolygon> ContourPolygonList, DfsuFile dfsuFile, int ParamCount, string ProvInit)
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

            string Date_UTC_TextStart = "";
            string Date_UTC_TextEnd = "";
            if (ProvInit == "NL")
            {
                TimeZoneInfo tst = TimeZoneInfo.FindSystemTimeZoneById("Newfoundland Standard Time");
                if (tst.IsDaylightSavingTime(dfsuFile.StartDateTime))
                {
                    Date_UTC_TextStart = dfsuFile.StartDateTime.AddSeconds(ParamCount * dfsuFile.TimeStepInSeconds).AddHours(3).AddMinutes(30).ToString("yyyy-MM-ddTHH:mm:ss");
                    Date_UTC_TextEnd = dfsuFile.StartDateTime.AddSeconds((ParamCount + 1) * dfsuFile.TimeStepInSeconds).AddHours(3).AddMinutes(30).ToString("yyyy-MM-ddTHH:mm:ss");
                }
                else
                {
                    Date_UTC_TextStart = dfsuFile.StartDateTime.AddSeconds(ParamCount * dfsuFile.TimeStepInSeconds).AddHours(2).AddMinutes(30).ToString("yyyy-MM-ddTHH:mm:ss");
                    Date_UTC_TextEnd = dfsuFile.StartDateTime.AddSeconds((ParamCount + 1) * dfsuFile.TimeStepInSeconds).AddHours(2).AddMinutes(30).ToString("yyyy-MM-ddTHH:mm:ss");
                }
            }
            else if (ProvInit == "QC")
            {
                TimeZoneInfo tst = TimeZoneInfo.FindSystemTimeZoneById("Eastern Standard Time");
                if (tst.IsDaylightSavingTime(dfsuFile.StartDateTime))
                {
                    Date_UTC_TextStart = dfsuFile.StartDateTime.AddSeconds(ParamCount * dfsuFile.TimeStepInSeconds).AddHours(5).ToString("yyyy-MM-ddTHH:mm:ss");
                    Date_UTC_TextEnd = dfsuFile.StartDateTime.AddSeconds((ParamCount + 1) * dfsuFile.TimeStepInSeconds).AddHours(5).ToString("yyyy-MM-ddTHH:mm:ss");
                }
                else
                {
                    Date_UTC_TextStart = dfsuFile.StartDateTime.AddSeconds(ParamCount * dfsuFile.TimeStepInSeconds).AddHours(4).ToString("yyyy-MM-ddTHH:mm:ss");
                    Date_UTC_TextEnd = dfsuFile.StartDateTime.AddSeconds((ParamCount + 1) * dfsuFile.TimeStepInSeconds).AddHours(4).ToString("yyyy-MM-ddTHH:mm:ss");
                }
            }
            else if (ProvInit == "BC")
            {
                TimeZoneInfo tst = TimeZoneInfo.FindSystemTimeZoneById("Pacific Standard Time");
                if (tst.IsDaylightSavingTime(dfsuFile.StartDateTime))
                {
                    Date_UTC_TextStart = dfsuFile.StartDateTime.AddSeconds(ParamCount * dfsuFile.TimeStepInSeconds).AddHours(8).ToString("yyyy-MM-ddTHH:mm:ss");
                    Date_UTC_TextEnd = dfsuFile.StartDateTime.AddSeconds((ParamCount + 1) * dfsuFile.TimeStepInSeconds).AddHours(8).ToString("yyyy-MM-ddTHH:mm:ss");
                }
                else
                {
                    Date_UTC_TextStart = dfsuFile.StartDateTime.AddSeconds(ParamCount * dfsuFile.TimeStepInSeconds).AddHours(7).ToString("yyyy-MM-ddTHH:mm:ss");
                    Date_UTC_TextEnd = dfsuFile.StartDateTime.AddSeconds((ParamCount + 1) * dfsuFile.TimeStepInSeconds).AddHours(7).ToString("yyyy-MM-ddTHH:mm:ss");
                }
            }
            else
            {
                TimeZoneInfo tst = TimeZoneInfo.FindSystemTimeZoneById("Atlantic Standard Time");
                if (tst.IsDaylightSavingTime(dfsuFile.StartDateTime))
                {
                    Date_UTC_TextStart = dfsuFile.StartDateTime.AddSeconds(ParamCount * dfsuFile.TimeStepInSeconds).AddHours(4).ToString("yyyy-MM-ddTHH:mm:ss");
                    Date_UTC_TextEnd = dfsuFile.StartDateTime.AddSeconds((ParamCount + 1) * dfsuFile.TimeStepInSeconds).AddHours(4).ToString("yyyy-MM-ddTHH:mm:ss");
                }
                else
                {
                    Date_UTC_TextStart = dfsuFile.StartDateTime.AddSeconds(ParamCount * dfsuFile.TimeStepInSeconds).AddHours(3).ToString("yyyy-MM-ddTHH:mm:ss");
                    Date_UTC_TextEnd = dfsuFile.StartDateTime.AddSeconds((ParamCount + 1) * dfsuFile.TimeStepInSeconds).AddHours(3).ToString("yyyy-MM-ddTHH:mm:ss");
                }
            }

            sb.AppendLine($@"    <begin>{Date_UTC_TextStart}</begin>");
            sb.AppendLine($@"    <end>{Date_UTC_TextEnd}</end>");
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
        private bool DrawKMLCurrentsStyle(StringBuilder sbStyleCurrentAnim)
        {
            sbStyleCurrentAnim.AppendLine(@"<Style id=""pink"">");
            sbStyleCurrentAnim.AppendLine(@"<LineStyle>");
            sbStyleCurrentAnim.AppendLine(@"<color>ffff00ff</color>");
            //sbStyleCurrentAnim.AppendLine(@"<gx:physicalWidth>12</gx:physicalWidth>");
            sbStyleCurrentAnim.AppendLine(@"</LineStyle>");
            sbStyleCurrentAnim.AppendLine(@"</Style>");

            sbStyleCurrentAnim.AppendLine(@"<Style id=""green"">");
            sbStyleCurrentAnim.AppendLine(@"<LineStyle>");
            sbStyleCurrentAnim.AppendLine(@"<color>ff00ff00</color>");
            //sbStyleCurrentAnim.AppendLine(@"<gx:physicalWidth>12</gx:physicalWidth>");
            sbStyleCurrentAnim.AppendLine(@"</LineStyle>");
            sbStyleCurrentAnim.AppendLine(@"</Style>");

            sbStyleCurrentAnim.AppendLine(@"<Style id=""yellow"">");
            sbStyleCurrentAnim.AppendLine(@"<LineStyle>");
            sbStyleCurrentAnim.AppendLine(@"<color>ff00ffff</color>");
            //sbStyleCurrentAnim.AppendLine(@"<gx:physicalWidth>12</gx:physicalWidth>");
            sbStyleCurrentAnim.AppendLine(@"</LineStyle>");
            sbStyleCurrentAnim.AppendLine(@"</Style>");

            return true;
        }
        private bool DrawKMLEstimatedDroguePathsStyle(StringBuilder sbStyleCurrentAnim)
        {
            // Style for blue circle
            sb.AppendLine($@"	<StyleMap id=""msn_placemark_circle_blue"">");
            sb.AppendLine($@"		<Pair>");
            sb.AppendLine($@"			<key>normal</key>");
            sb.AppendLine($@"			<styleUrl>#sn_placemark_circle_blue</styleUrl>");
            sb.AppendLine($@"		</Pair>");
            sb.AppendLine($@"		<Pair>");
            sb.AppendLine($@"			<key>highlight</key>");
            sb.AppendLine($@"			<styleUrl>#sh_placemark_circle_highlight_blue</styleUrl>");
            sb.AppendLine($@"		</Pair>");
            sb.AppendLine($@"	</StyleMap>");
            sb.AppendLine($@"	<Style id=""sh_placemark_circle_highlight_blue"">");
            sb.AppendLine($@"		<IconStyle>");
            sb.AppendLine($@"			<color>ffff0000</color>");
            sb.AppendLine($@"			<scale>0.6</scale>");
            sb.AppendLine($@"			<Icon>");
            sb.AppendLine($@"				<href>http://maps.google.com/mapfiles/kml/shapes/placemark_circle_highlight.png</href>");
            sb.AppendLine($@"			</Icon>");
            sb.AppendLine($@"		</IconStyle>");
            sb.AppendLine($@"		<LineStyle>");
            sb.AppendLine($@"			<color>ffff0000</color>");
            sb.AppendLine($@"		</LineStyle>");
            sb.AppendLine($@"	</Style>");
            sb.AppendLine($@"	<Style id=""sn_placemark_circle_blue"">");
            sb.AppendLine($@"		<IconStyle>");
            sb.AppendLine($@"			<color>ffff0000</color>");
            sb.AppendLine($@"			<scale>0.6</scale>");
            sb.AppendLine($@"			<Icon>");
            sb.AppendLine($@"				<href>http://maps.google.com/mapfiles/kml/shapes/placemark_circle.png</href>");
            sb.AppendLine($@"			</Icon>");
            sb.AppendLine($@"		</IconStyle>");
            sb.AppendLine($@"		<LineStyle>");
            sb.AppendLine($@"			<color>ffff0000</color>");
            sb.AppendLine($@"		</LineStyle>");
            sb.AppendLine($@" </Style>");

            // Style for green circle
            sb.AppendLine($@"	<StyleMap id=""msn_placemark_circle_green"">");
            sb.AppendLine($@"		<Pair>");
            sb.AppendLine($@"			<key>normal</key>");
            sb.AppendLine($@"			<styleUrl>#sn_placemark_circle_green</styleUrl>");
            sb.AppendLine($@"		</Pair>");
            sb.AppendLine($@"		<Pair>");
            sb.AppendLine($@"			<key>highlight</key>");
            sb.AppendLine($@"			<styleUrl>#sh_placemark_circle_highlight_green</styleUrl>");
            sb.AppendLine($@"		</Pair>");
            sb.AppendLine($@"	</StyleMap>");
            sb.AppendLine($@"	<Style id=""sh_placemark_circle_highlight_green"">");
            sb.AppendLine($@"		<IconStyle>");
            sb.AppendLine($@"			<color>ff00ff00</color>");
            sb.AppendLine($@"			<scale>0.6</scale>");
            sb.AppendLine($@"			<Icon>");
            sb.AppendLine($@"				<href>http://maps.google.com/mapfiles/kml/shapes/placemark_circle_highlight.png</href>");
            sb.AppendLine($@"			</Icon>");
            sb.AppendLine($@"		</IconStyle>");
            sb.AppendLine($@"		<LineStyle>");
            sb.AppendLine($@"			<color>ff00ff00</color>");
            sb.AppendLine($@"		</LineStyle>");
            sb.AppendLine($@"	</Style>");
            sb.AppendLine($@"	<Style id=""sn_placemark_circle_green"">");
            sb.AppendLine($@"		<IconStyle>");
            sb.AppendLine($@"			<color>ff00ff00</color>");
            sb.AppendLine($@"			<scale>0.6</scale>");
            sb.AppendLine($@"			<Icon>");
            sb.AppendLine($@"				<href>http://maps.google.com/mapfiles/kml/shapes/placemark_circle.png</href>");
            sb.AppendLine($@"			</Icon>");
            sb.AppendLine($@"		</IconStyle>");
            sb.AppendLine($@"		<LineStyle>");
            sb.AppendLine($@"			<color>ff00ff00</color>");
            sb.AppendLine($@"		</LineStyle>");
            sb.AppendLine($@" </Style>");

            // Style for red circle
            sb.AppendLine($@"	<StyleMap id=""msn_placemark_circle_red"">");
            sb.AppendLine($@"		<Pair>");
            sb.AppendLine($@"			<key>normal</key>");
            sb.AppendLine($@"			<styleUrl>#sn_placemark_circle_red</styleUrl>");
            sb.AppendLine($@"		</Pair>");
            sb.AppendLine($@"		<Pair>");
            sb.AppendLine($@"			<key>highlight</key>");
            sb.AppendLine($@"			<styleUrl>#sh_placemark_circle_highlight_red</styleUrl>");
            sb.AppendLine($@"		</Pair>");
            sb.AppendLine($@"	</StyleMap>");
            sb.AppendLine($@"	<Style id=""sh_placemark_circle_highlight_red"">");
            sb.AppendLine($@"		<IconStyle>");
            sb.AppendLine($@"			<color>ff0000ff</color>");
            sb.AppendLine($@"			<scale>0.6</scale>");
            sb.AppendLine($@"			<Icon>");
            sb.AppendLine($@"				<href>http://maps.google.com/mapfiles/kml/shapes/placemark_circle_highlight.png</href>");
            sb.AppendLine($@"			</Icon>");
            sb.AppendLine($@"		</IconStyle>");
            sb.AppendLine($@"		<LineStyle>");
            sb.AppendLine($@"			<color>ff0000ff</color>");
            sb.AppendLine($@"		</LineStyle>");
            sb.AppendLine($@"	</Style>");
            sb.AppendLine($@"	<Style id=""sn_placemark_circle_red"">");
            sb.AppendLine($@"		<IconStyle>");
            sb.AppendLine($@"			<color>ff0000ff</color>");
            sb.AppendLine($@"			<scale>0.6</scale>");
            sb.AppendLine($@"			<Icon>");
            sb.AppendLine($@"				<href>http://maps.google.com/mapfiles/kml/shapes/placemark_circle.png</href>");
            sb.AppendLine($@"			</Icon>");
            sb.AppendLine($@"		</IconStyle>");
            sb.AppendLine($@"		<LineStyle>");
            sb.AppendLine($@"			<color>ff0000ff</color>");
            sb.AppendLine($@"		</LineStyle>");
            sb.AppendLine($@" </Style>");

            // Style for purple circle
            sb.AppendLine($@"	<StyleMap id=""msn_placemark_circle_purple"">");
            sb.AppendLine($@"		<Pair>");
            sb.AppendLine($@"			<key>normal</key>");
            sb.AppendLine($@"			<styleUrl>#sn_placemark_circle_purple</styleUrl>");
            sb.AppendLine($@"		</Pair>");
            sb.AppendLine($@"		<Pair>");
            sb.AppendLine($@"			<key>highlight</key>");
            sb.AppendLine($@"			<styleUrl>#sh_placemark_circle_highlight_purple</styleUrl>");
            sb.AppendLine($@"		</Pair>");
            sb.AppendLine($@"	</StyleMap>");
            sb.AppendLine($@"	<Style id=""sh_placemark_circle_highlight_purple"">");
            sb.AppendLine($@"		<IconStyle>");
            sb.AppendLine($@"			<color>ff800080</color>");
            sb.AppendLine($@"			<scale>0.6</scale>");
            sb.AppendLine($@"			<Icon>");
            sb.AppendLine($@"				<href>http://maps.google.com/mapfiles/kml/shapes/placemark_circle_highlight.png</href>");
            sb.AppendLine($@"			</Icon>");
            sb.AppendLine($@"		</IconStyle>");
            sb.AppendLine($@"		<LineStyle>");
            sb.AppendLine($@"			<color>ff800080</color>");
            sb.AppendLine($@"		</LineStyle>");
            sb.AppendLine($@"	</Style>");
            sb.AppendLine($@"	<Style id=""sn_placemark_circle_purple"">");
            sb.AppendLine($@"		<IconStyle>");
            sb.AppendLine($@"			<color>ff800080</color>");
            sb.AppendLine($@"			<scale>0.6</scale>");
            sb.AppendLine($@"			<Icon>");
            sb.AppendLine($@"				<href>http://maps.google.com/mapfiles/kml/shapes/placemark_circle.png</href>");
            sb.AppendLine($@"			</Icon>");
            sb.AppendLine($@"		</IconStyle>");
            sb.AppendLine($@"		<LineStyle>");
            sb.AppendLine($@"			<color>ff800080</color>");
            sb.AppendLine($@"		</LineStyle>");
            sb.AppendLine($@" </Style>");

            // Style for blue square
            sb.AppendLine($@"	<StyleMap id=""msn_placemark_square_blue"">");
            sb.AppendLine($@"		<Pair>");
            sb.AppendLine($@"			<key>normal</key>");
            sb.AppendLine($@"			<styleUrl>#sn_placemark_square_blue</styleUrl>");
            sb.AppendLine($@"		</Pair>");
            sb.AppendLine($@"		<Pair>");
            sb.AppendLine($@"			<key>highlight</key>");
            sb.AppendLine($@"			<styleUrl>#sh_placemark_square_highlight_blue</styleUrl>");
            sb.AppendLine($@"		</Pair>");
            sb.AppendLine($@"	</StyleMap>");
            sb.AppendLine($@"	<Style id=""sh_placemark_square_highlight_blue"">");
            sb.AppendLine($@"		<IconStyle>");
            sb.AppendLine($@"			<color>ffff0000</color>");
            sb.AppendLine($@"			<scale>0.6</scale>");
            sb.AppendLine($@"			<Icon>");
            sb.AppendLine($@"				<href>http://maps.google.com/mapfiles/kml/shapes/placemark_square_highlight.png</href>");
            sb.AppendLine($@"			</Icon>");
            sb.AppendLine($@"		</IconStyle>");
            sb.AppendLine($@"		<LineStyle>");
            sb.AppendLine($@"			<color>ffff0000</color>");
            sb.AppendLine($@"		</LineStyle>");
            sb.AppendLine($@"	</Style>");
            sb.AppendLine($@"	<Style id=""sn_placemark_square_blue"">");
            sb.AppendLine($@"		<IconStyle>");
            sb.AppendLine($@"			<color>ffff0000</color>");
            sb.AppendLine($@"			<scale>0.6</scale>");
            sb.AppendLine($@"			<Icon>");
            sb.AppendLine($@"				<href>http://maps.google.com/mapfiles/kml/shapes/placemark_square.png</href>");
            sb.AppendLine($@"			</Icon>");
            sb.AppendLine($@"		</IconStyle>");
            sb.AppendLine($@"		<LineStyle>");
            sb.AppendLine($@"			<color>ffff0000</color>");
            sb.AppendLine($@"		</LineStyle>");
            sb.AppendLine($@" </Style>");

            // Style for green square
            sb.AppendLine($@"	<StyleMap id=""msn_placemark_square_green"">");
            sb.AppendLine($@"		<Pair>");
            sb.AppendLine($@"			<key>normal</key>");
            sb.AppendLine($@"			<styleUrl>#sn_placemark_square_green</styleUrl>");
            sb.AppendLine($@"		</Pair>");
            sb.AppendLine($@"		<Pair>");
            sb.AppendLine($@"			<key>highlight</key>");
            sb.AppendLine($@"			<styleUrl>#sh_placemark_square_highlight_green</styleUrl>");
            sb.AppendLine($@"		</Pair>");
            sb.AppendLine($@"	</StyleMap>");
            sb.AppendLine($@"	<Style id=""sh_placemark_square_highlight_green"">");
            sb.AppendLine($@"		<IconStyle>");
            sb.AppendLine($@"			<color>ff00ff00</color>");
            sb.AppendLine($@"			<scale>0.6</scale>");
            sb.AppendLine($@"			<Icon>");
            sb.AppendLine($@"				<href>http://maps.google.com/mapfiles/kml/shapes/placemark_square_highlight.png</href>");
            sb.AppendLine($@"			</Icon>");
            sb.AppendLine($@"		</IconStyle>");
            sb.AppendLine($@"		<LineStyle>");
            sb.AppendLine($@"			<color>ff00ff00</color>");
            sb.AppendLine($@"		</LineStyle>");
            sb.AppendLine($@"	</Style>");
            sb.AppendLine($@"	<Style id=""sn_placemark_square_green"">");
            sb.AppendLine($@"		<IconStyle>");
            sb.AppendLine($@"			<color>ff00ff00</color>");
            sb.AppendLine($@"			<scale>0.6</scale>");
            sb.AppendLine($@"			<Icon>");
            sb.AppendLine($@"				<href>http://maps.google.com/mapfiles/kml/shapes/placemark_square.png</href>");
            sb.AppendLine($@"			</Icon>");
            sb.AppendLine($@"		</IconStyle>");
            sb.AppendLine($@"		<LineStyle>");
            sb.AppendLine($@"			<color>ff00ff00</color>");
            sb.AppendLine($@"		</LineStyle>");
            sb.AppendLine($@" </Style>");

            // Style for red square
            sb.AppendLine($@"	<StyleMap id=""msn_placemark_square_red"">");
            sb.AppendLine($@"		<Pair>");
            sb.AppendLine($@"			<key>normal</key>");
            sb.AppendLine($@"			<styleUrl>#sn_placemark_square_red</styleUrl>");
            sb.AppendLine($@"		</Pair>");
            sb.AppendLine($@"		<Pair>");
            sb.AppendLine($@"			<key>highlight</key>");
            sb.AppendLine($@"			<styleUrl>#sh_placemark_square_highlight_red</styleUrl>");
            sb.AppendLine($@"		</Pair>");
            sb.AppendLine($@"	</StyleMap>");
            sb.AppendLine($@"	<Style id=""sh_placemark_square_highlight_red"">");
            sb.AppendLine($@"		<IconStyle>");
            sb.AppendLine($@"			<color>ff0000ff</color>");
            sb.AppendLine($@"			<scale>0.6</scale>");
            sb.AppendLine($@"			<Icon>");
            sb.AppendLine($@"				<href>http://maps.google.com/mapfiles/kml/shapes/placemark_square_highlight.png</href>");
            sb.AppendLine($@"			</Icon>");
            sb.AppendLine($@"		</IconStyle>");
            sb.AppendLine($@"		<LineStyle>");
            sb.AppendLine($@"			<color>ff0000ff</color>");
            sb.AppendLine($@"		</LineStyle>");
            sb.AppendLine($@"	</Style>");
            sb.AppendLine($@"	<Style id=""sn_placemark_square_red"">");
            sb.AppendLine($@"		<IconStyle>");
            sb.AppendLine($@"			<color>ff0000ff</color>");
            sb.AppendLine($@"			<scale>0.6</scale>");
            sb.AppendLine($@"			<Icon>");
            sb.AppendLine($@"				<href>http://maps.google.com/mapfiles/kml/shapes/placemark_square.png</href>");
            sb.AppendLine($@"			</Icon>");
            sb.AppendLine($@"		</IconStyle>");
            sb.AppendLine($@"		<LineStyle>");
            sb.AppendLine($@"			<color>ff0000ff</color>");
            sb.AppendLine($@"		</LineStyle>");
            sb.AppendLine($@" </Style>");

            // Style for purple square
            sb.AppendLine($@"	<StyleMap id=""msn_placemark_square_purple"">");
            sb.AppendLine($@"		<Pair>");
            sb.AppendLine($@"			<key>normal</key>");
            sb.AppendLine($@"			<styleUrl>#sn_placemark_square_purple</styleUrl>");
            sb.AppendLine($@"		</Pair>");
            sb.AppendLine($@"		<Pair>");
            sb.AppendLine($@"			<key>highlight</key>");
            sb.AppendLine($@"			<styleUrl>#sh_placemark_square_highlight_purple</styleUrl>");
            sb.AppendLine($@"		</Pair>");
            sb.AppendLine($@"	</StyleMap>");
            sb.AppendLine($@"	<Style id=""sh_placemark_square_highlight_purple"">");
            sb.AppendLine($@"		<IconStyle>");
            sb.AppendLine($@"			<color>ff800080</color>");
            sb.AppendLine($@"			<scale>0.6</scale>");
            sb.AppendLine($@"			<Icon>");
            sb.AppendLine($@"				<href>http://maps.google.com/mapfiles/kml/shapes/placemark_square_highlight.png</href>");
            sb.AppendLine($@"			</Icon>");
            sb.AppendLine($@"		</IconStyle>");
            sb.AppendLine($@"		<LineStyle>");
            sb.AppendLine($@"			<color>ff800080</color>");
            sb.AppendLine($@"		</LineStyle>");
            sb.AppendLine($@"	</Style>");
            sb.AppendLine($@"	<Style id=""sn_placemark_square_purple"">");
            sb.AppendLine($@"		<IconStyle>");
            sb.AppendLine($@"			<color>ff800080</color>");
            sb.AppendLine($@"			<scale>0.6</scale>");
            sb.AppendLine($@"			<Icon>");
            sb.AppendLine($@"				<href>http://maps.google.com/mapfiles/kml/shapes/placemark_square.png</href>");
            sb.AppendLine($@"			</Icon>");
            sb.AppendLine($@"		</IconStyle>");
            sb.AppendLine($@"		<LineStyle>");
            sb.AppendLine($@"			<color>ff800080</color>");
            sb.AppendLine($@"		</LineStyle>");
            sb.AppendLine($@" </Style>");

            // Style for blue shaded_dot
            sb.AppendLine($@"	<StyleMap id=""msn_placemark_shaded_dot_blue"">");
            sb.AppendLine($@"		<Pair>");
            sb.AppendLine($@"			<key>normal</key>");
            sb.AppendLine($@"			<styleUrl>#sn_placemark_shaded_dot_blue</styleUrl>");
            sb.AppendLine($@"		</Pair>");
            sb.AppendLine($@"		<Pair>");
            sb.AppendLine($@"			<key>highlight</key>");
            sb.AppendLine($@"			<styleUrl>#sh_placemark_shaded_dot_highlight_blue</styleUrl>");
            sb.AppendLine($@"		</Pair>");
            sb.AppendLine($@"	</StyleMap>");
            sb.AppendLine($@"	<Style id=""sh_placemark_shaded_dot_highlight_blue"">");
            sb.AppendLine($@"		<IconStyle>");
            sb.AppendLine($@"			<color>ffff0000</color>");
            sb.AppendLine($@"			<scale>0.6</scale>");
            sb.AppendLine($@"			<Icon>");
            sb.AppendLine($@"				<href>http://maps.google.com/mapfiles/kml/shapes/shaded_dot.png</href>");
            sb.AppendLine($@"			</Icon>");
            sb.AppendLine($@"		</IconStyle>");
            sb.AppendLine($@"		<LineStyle>");
            sb.AppendLine($@"			<color>ffff0000</color>");
            sb.AppendLine($@"		</LineStyle>");
            sb.AppendLine($@"	</Style>");
            sb.AppendLine($@"	<Style id=""sn_placemark_shaded_dot_blue"">");
            sb.AppendLine($@"		<IconStyle>");
            sb.AppendLine($@"			<color>ffff0000</color>");
            sb.AppendLine($@"			<scale>0.6</scale>");
            sb.AppendLine($@"			<Icon>");
            sb.AppendLine($@"				<href>http://maps.google.com/mapfiles/kml/shapes/shaded_dot.png</href>");
            sb.AppendLine($@"			</Icon>");
            sb.AppendLine($@"		</IconStyle>");
            sb.AppendLine($@"		<LineStyle>");
            sb.AppendLine($@"			<color>ffff0000</color>");
            sb.AppendLine($@"		</LineStyle>");
            sb.AppendLine($@" </Style>");

            // Style for green shaded_dot
            sb.AppendLine($@"	<StyleMap id=""msn_placemark_shaded_dot_green"">");
            sb.AppendLine($@"		<Pair>");
            sb.AppendLine($@"			<key>normal</key>");
            sb.AppendLine($@"			<styleUrl>#sn_placemark_shaded_dot_green</styleUrl>");
            sb.AppendLine($@"		</Pair>");
            sb.AppendLine($@"		<Pair>");
            sb.AppendLine($@"			<key>highlight</key>");
            sb.AppendLine($@"			<styleUrl>#sh_placemark_shaded_dot_highlight_green</styleUrl>");
            sb.AppendLine($@"		</Pair>");
            sb.AppendLine($@"	</StyleMap>");
            sb.AppendLine($@"	<Style id=""sh_placemark_shaded_dot_highlight_green"">");
            sb.AppendLine($@"		<IconStyle>");
            sb.AppendLine($@"			<color>ff00ff00</color>");
            sb.AppendLine($@"			<scale>0.6</scale>");
            sb.AppendLine($@"			<Icon>");
            sb.AppendLine($@"				<href>http://maps.google.com/mapfiles/kml/shapes/shaded_dot.png</href>");
            sb.AppendLine($@"			</Icon>");
            sb.AppendLine($@"		</IconStyle>");
            sb.AppendLine($@"		<LineStyle>");
            sb.AppendLine($@"			<color>ff00ff00</color>");
            sb.AppendLine($@"		</LineStyle>");
            sb.AppendLine($@"	</Style>");
            sb.AppendLine($@"	<Style id=""sn_placemark_shaded_dot_green"">");
            sb.AppendLine($@"		<IconStyle>");
            sb.AppendLine($@"			<color>ff00ff00</color>");
            sb.AppendLine($@"			<scale>0.6</scale>");
            sb.AppendLine($@"			<Icon>");
            sb.AppendLine($@"				<href>http://maps.google.com/mapfiles/kml/shapes/shaded_dot.png</href>");
            sb.AppendLine($@"			</Icon>");
            sb.AppendLine($@"		</IconStyle>");
            sb.AppendLine($@"		<LineStyle>");
            sb.AppendLine($@"			<color>ff00ff00</color>");
            sb.AppendLine($@"		</LineStyle>");
            sb.AppendLine($@" </Style>");

            // Style for red shaded_dot
            sb.AppendLine($@"	<StyleMap id=""msn_placemark_shaded_dot_red"">");
            sb.AppendLine($@"		<Pair>");
            sb.AppendLine($@"			<key>normal</key>");
            sb.AppendLine($@"			<styleUrl>#sn_placemark_shaded_dot_red</styleUrl>");
            sb.AppendLine($@"		</Pair>");
            sb.AppendLine($@"		<Pair>");
            sb.AppendLine($@"			<key>highlight</key>");
            sb.AppendLine($@"			<styleUrl>#sh_placemark_shaded_dot_highlight_red</styleUrl>");
            sb.AppendLine($@"		</Pair>");
            sb.AppendLine($@"	</StyleMap>");
            sb.AppendLine($@"	<Style id=""sh_placemark_shaded_dot_highlight_red"">");
            sb.AppendLine($@"		<IconStyle>");
            sb.AppendLine($@"			<color>ff0000ff</color>");
            sb.AppendLine($@"			<scale>0.6</scale>");
            sb.AppendLine($@"			<Icon>");
            sb.AppendLine($@"				<href>http://maps.google.com/mapfiles/kml/shapes/shaded_dot.png</href>");
            sb.AppendLine($@"			</Icon>");
            sb.AppendLine($@"		</IconStyle>");
            sb.AppendLine($@"		<LineStyle>");
            sb.AppendLine($@"			<color>ff0000ff</color>");
            sb.AppendLine($@"		</LineStyle>");
            sb.AppendLine($@"	</Style>");
            sb.AppendLine($@"	<Style id=""sn_placemark_shaded_dot_red"">");
            sb.AppendLine($@"		<IconStyle>");
            sb.AppendLine($@"			<color>ff0000ff</color>");
            sb.AppendLine($@"			<scale>0.6</scale>");
            sb.AppendLine($@"			<Icon>");
            sb.AppendLine($@"				<href>http://maps.google.com/mapfiles/kml/shapes/shaded_dot.png</href>");
            sb.AppendLine($@"			</Icon>");
            sb.AppendLine($@"		</IconStyle>");
            sb.AppendLine($@"		<LineStyle>");
            sb.AppendLine($@"			<color>ff0000ff</color>");
            sb.AppendLine($@"		</LineStyle>");
            sb.AppendLine($@" </Style>");

            // Style for purple shaded_dot
            sb.AppendLine($@"	<StyleMap id=""msn_placemark_shaded_dot_purple"">");
            sb.AppendLine($@"		<Pair>");
            sb.AppendLine($@"			<key>normal</key>");
            sb.AppendLine($@"			<styleUrl>#sn_placemark_shaded_dot_purple</styleUrl>");
            sb.AppendLine($@"		</Pair>");
            sb.AppendLine($@"		<Pair>");
            sb.AppendLine($@"			<key>highlight</key>");
            sb.AppendLine($@"			<styleUrl>#sh_placemark_shaded_dot_highlight_purple</styleUrl>");
            sb.AppendLine($@"		</Pair>");
            sb.AppendLine($@"	</StyleMap>");
            sb.AppendLine($@"	<Style id=""sh_placemark_shaded_dot_highlight_purple"">");
            sb.AppendLine($@"		<IconStyle>");
            sb.AppendLine($@"			<color>ff800080</color>");
            sb.AppendLine($@"			<scale>0.6</scale>");
            sb.AppendLine($@"			<Icon>");
            sb.AppendLine($@"				<href>http://maps.google.com/mapfiles/kml/shapes/shaded_dot.png</href>");
            sb.AppendLine($@"			</Icon>");
            sb.AppendLine($@"		</IconStyle>");
            sb.AppendLine($@"		<LineStyle>");
            sb.AppendLine($@"			<color>ff800080</color>");
            sb.AppendLine($@"		</LineStyle>");
            sb.AppendLine($@"	</Style>");
            sb.AppendLine($@"	<Style id=""sn_placemark_shaded_dot_purple"">");
            sb.AppendLine($@"		<IconStyle>");
            sb.AppendLine($@"			<color>ff800080</color>");
            sb.AppendLine($@"			<scale>0.6</scale>");
            sb.AppendLine($@"			<Icon>");
            sb.AppendLine($@"				<href>http://maps.google.com/mapfiles/kml/shapes/shaded_dot.png</href>");
            sb.AppendLine($@"			</Icon>");
            sb.AppendLine($@"		</IconStyle>");
            sb.AppendLine($@"		<LineStyle>");
            sb.AppendLine($@"			<color>ff800080</color>");
            sb.AppendLine($@"		</LineStyle>");
            sb.AppendLine($@" </Style>");


            return true;
        }
        public bool FillElementLayerList(DfsuFile dfsuFile, List<Element> ElementList, List<ElementLayer> ElementLayerList, List<NodeLayer> topNodeLayerList, List<NodeLayer> bottomNodeLayerList)
        {
            string NotUsed = "";

            if (dfsuFile.DfsuFileType == DfsuFileType.Dfsu3DSigmaZ)
            {
                NotUsed = string.Format(TaskRunnerServiceRes._NotImplemented, "Z Level");
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_NotImplemented", "Z Level");
                return false;
            }

            if (dfsuFile.DfsuFileType == DfsuFileType.Dfsu3DSigma)
            {
                // doing type 32
                var coordList = (from el in ElementList
                                 where el.Type == 32
                                 select new { X1 = el.NodeList[0].X, X2 = el.NodeList[1].X, X3 = el.NodeList[2].X }).Distinct().ToList();

                foreach (var coord in coordList)
                {
                    int Layer = 1;
                    List<Element> ColumnElementList = (from el1 in ElementList
                                                       where el1.Type == 32
                                                       && el1.NodeList[0].X == coord.X1
                                                       && el1.NodeList[1].X == coord.X2
                                                       && el1.NodeList[2].X == coord.X3
                                                       orderby dfsuFile.Z[el1.NodeList[0].ID - 1]
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

                        topNodeLayerList.Add(nl0);
                        topNodeLayerList.Add(nl1);
                        topNodeLayerList.Add(nl2);

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

                        bottomNodeLayerList.Add(nl3);
                        bottomNodeLayerList.Add(nl4);
                        bottomNodeLayerList.Add(nl5);

                        Layer += 1;
                    }

                }

                var coordList2 = (from el in ElementList
                                  where el.Type == 33
                                  select new { X1 = el.NodeList[0].X, X2 = el.NodeList[1].X, X3 = el.NodeList[2].X, X4 = el.NodeList[3].X }).Distinct().ToList();

                foreach (var coord in coordList2)
                {
                    int Layer = 1;
                    List<Element> ColumnElementList = (from el1 in ElementList
                                                       where el1.Type == 33
                                                       && el1.NodeList[0].X == coord.X1
                                                       && el1.NodeList[1].X == coord.X2
                                                       && el1.NodeList[2].X == coord.X3
                                                       && el1.NodeList[3].X == coord.X4
                                                       orderby dfsuFile.Z[el1.NodeList[0].ID - 1]
                                                       select el1).ToList<Element>();

                    for (int j = 0; j < ColumnElementList.Count; j++)
                    {
                        ElementLayer elementLayer = new ElementLayer();
                        elementLayer.Layer = Layer;
                        elementLayer.ZMin = (from nz in ColumnElementList[j].NodeList select dfsuFile.Z[nz.ID - 1]).Min();
                        elementLayer.ZMax = (from nz in ColumnElementList[j].NodeList select dfsuFile.Z[nz.ID - 1]).Max();
                        elementLayer.Element = ColumnElementList[j];
                        ElementLayerList.Add(elementLayer);

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

                        NodeLayer nl3 = new NodeLayer();
                        nl3.Layer = Layer;
                        nl3.Z = dfsuFile.Z[ColumnElementList[j].NodeList[3].ID - 1];
                        nl3.Node = ColumnElementList[j].NodeList[3];

                        topNodeLayerList.Add(nl0);
                        topNodeLayerList.Add(nl1);
                        topNodeLayerList.Add(nl2);
                        topNodeLayerList.Add(nl3);

                        NodeLayer nl4 = new NodeLayer();
                        nl4.Layer = Layer;
                        nl4.Z = 0;
                        nl4.Node = ColumnElementList[j].NodeList[4];

                        NodeLayer nl5 = new NodeLayer();
                        nl5.Layer = Layer;
                        nl5.Z = 0;
                        nl5.Node = ColumnElementList[j].NodeList[5];

                        NodeLayer nl6 = new NodeLayer();
                        nl6.Layer = Layer;
                        nl6.Z = 0;
                        nl6.Node = ColumnElementList[j].NodeList[6];

                        NodeLayer nl7 = new NodeLayer();
                        nl7.Layer = Layer;
                        nl7.Z = 0;
                        nl7.Node = ColumnElementList[j].NodeList[7];


                        bottomNodeLayerList.Add(nl4);
                        bottomNodeLayerList.Add(nl5);
                        bottomNodeLayerList.Add(nl6);
                        bottomNodeLayerList.Add(nl7);

                        Layer += 1;
                    }

                }


                //List<ElementLayer> TempElementLayerList = (from el in ElementLayerList
                //                                           orderby el.Element.ID
                //                                           select el).Distinct().ToList();

                ////ElementLayerList = new List<ElementLayer>();
                //int OldElemID = 0;
                //foreach (ElementLayer el in TempElementLayerList)
                //{
                //    if (OldElemID == el.Element.ID)
                //    {
                //        ElementLayerList.Remove(el);
                //    }
                //    OldElemID = el.Element.ID;
                //}

                //List<NodeLayer> TempNodeLayerList = (from nl in topNodeLayerList
                //                                     orderby nl.Node.ID
                //                                     select nl).Distinct().ToList();

                //topNodeLayerList = new List<NodeLayer>();
                //int OldID = 0;
                //foreach (NodeLayer nl in TempNodeLayerList)
                //{
                //    if (OldID != nl.Node.ID)
                //    {
                //        topNodeLayerList.Add(nl);
                //        OldID = nl.Node.ID;
                //    }
                //}

                //if (bottomNodeLayerList.Count() > 0)
                //{
                //    TempNodeLayerList = (from nl in bottomNodeLayerList
                //                         orderby nl.Node.ID
                //                         select nl).Distinct().ToList();

                //    bottomNodeLayerList = new List<NodeLayer>();
                //    OldID = 0;
                //    foreach (NodeLayer nl in TempNodeLayerList)
                //    {
                //        if (OldID != nl.Node.ID)
                //        {
                //            bottomNodeLayerList.Add(nl);
                //            OldID = nl.Node.ID;
                //        }
                //    }
                //}

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
        public bool FillRequiredList(DfsuFile dfsuFile, List<Element> elementList, List<ElementLayer> elementLayerList, List<Node> nodeList, List<NodeLayer> topNodeLayerList, List<NodeLayer> bottomNodeLayerList)
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
        private bool FillVectors21_32(Element el, List<Element> UniqueElementList, float ContourValue)
        {
            string NotUsed = "";

            Node Node0 = el.NodeList[0];
            Node Node1 = el.NodeList[1];
            Node Node2 = el.NodeList[2];

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
        private bool FillVectors25_33(Element el, List<Element> UniqueElementList, float ContourValue)
        {
            string NotUsed = "";

            Node Node0 = el.NodeList[0];
            Node Node1 = el.NodeList[1];
            Node Node2 = el.NodeList[2];
            Node Node3 = el.NodeList[3];

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
        public List<ElementLayer> GetElementSurrondingEachPoint(List<ElementLayer> ElementLayerList, List<Node> Nodes)
        {
            List<ElementLayer> AllElementList = new List<ElementLayer>();

            foreach (ElementLayer el in ElementLayerList)
            {
                float XMin = (from a in el.Element.NodeList
                              select a.X).Min();
                float YMin = (from a in el.Element.NodeList
                              select a.Y).Min();
                float XMax = (from a in el.Element.NodeList
                              select a.X).Max();
                float YMax = (from a in el.Element.NodeList
                              select a.Y).Max();

                foreach (Node n in Nodes)
                {
                    if ((n.X > XMin && n.X < XMax) && (n.Y > YMin && n.Y < YMax))
                    {
                        Point p = new Point((int)(n.X * 10000000), (int)(n.Y * 10000000));
                        if (el.Element.Type == 21 || el.Element.Type == 32)
                        {
                            Point[] poly =
                            {
                                new Point() { X = (int)(el.Element.NodeList[0].X*10000000), Y = (int)(el.Element.NodeList[0].Y*10000000) },
                                new Point() { X = (int)(el.Element.NodeList[1].X*10000000), Y = (int)(el.Element.NodeList[1].Y*10000000) },
                                new Point() { X = (int)(el.Element.NodeList[2].X*10000000), Y = (int)(el.Element.NodeList[2].Y*10000000) },
                                new Point() { X = (int)(el.Element.NodeList[0].X*10000000), Y = (int)(el.Element.NodeList[0].Y*10000000) },
                                           };

                            if (PointInPolygon(p, poly))
                            {
                                AllElementList.Add(el);
                            }
                        }
                        else if (el.Element.Type == 25 || el.Element.Type == 33)
                        {
                            Point[] poly =
                            {
                                new Point() { X = (int)(el.Element.NodeList[0].X*10000000), Y = (int)(el.Element.NodeList[0].Y*10000000) },
                                new Point() { X = (int)(el.Element.NodeList[1].X*10000000), Y = (int)(el.Element.NodeList[1].Y*10000000) },
                                new Point() { X = (int)(el.Element.NodeList[2].X*10000000), Y = (int)(el.Element.NodeList[2].Y*10000000) },
                                new Point() { X = (int)(el.Element.NodeList[3].X*10000000), Y = (int)(el.Element.NodeList[3].Y*10000000) },
                                new Point() { X = (int)(el.Element.NodeList[0].X*10000000), Y = (int)(el.Element.NodeList[0].Y*10000000) },
                                           };
                            if (PointInPolygon(p, poly))
                            {
                                AllElementList.Add(el);
                            }
                        }
                        else
                        {
                            return null;
                        }
                    }
                }
            }

            return AllElementList;
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
                        FileName = keyword.GetParameter(1).ToResultFileName();
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
        public string GetParameterResultFileName(PFSFile pfsFile, string Path, string Keyword)
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
                    FileName = keyword.GetParameter(1).ToResultFileName();
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
            string HydroFileName = GetParameterResultFileName(pfsFile, "FemEngineHD/HYDRODYNAMIC_MODULE/OUTPUTS/OUTPUT_1", "file_name");
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
        public DfsuFile GetTransportDfsuFile()
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
            string TransFileName = GetParameterResultFileName(pfsFile, "FemEngineHD/TRANSPORT_MODULE/OUTPUTS/OUTPUT_1", "file_name");
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
        private List<ElementLayer> ParseKMLPath(string KMLTextPathForVector, List<ElementLayer> ElementLayerList)
        {
            string NotUsed = "";

            List<ElementLayer> AllElementList = new List<ElementLayer>();
            List<Node> PathNodeList = new List<Node>();

            if (KMLTextPathForVector.Trim() == "")
            {
                foreach (ElementLayer el in ElementLayerList)
                {
                    AllElementList.Add(el);
                }
            }
            else
            {
                try
                {
                    XmlReader reader = XmlReader.Create(new StringReader(KMLTextPathForVector));
                    while (reader.Read())
                    {
                        if (reader.Name == "coordinates")
                        {
                            string AllCoordinates = reader.ReadElementContentAsString().Trim();

                            string[] xyzArray = AllCoordinates.Split(" ".ToCharArray()[0]);
                            foreach (string xyz in xyzArray)
                            {
                                string[] xyzStr = xyz.Split(",".ToCharArray()[0]);
                                if (xyzStr.Count() != 3)
                                {
                                    return null;
                                }
                                Node n = new Node();
                                if (Thread.CurrentThread.CurrentCulture.Name == "fr-CA")
                                {
                                    n.X = float.Parse(xyzStr[0].Replace(".", ","));
                                    n.Y = float.Parse(xyzStr[1].Replace(".", ","));
                                }
                                else
                                {
                                    n.X = float.Parse(xyzStr[0]);
                                    n.Y = float.Parse(xyzStr[1]);
                                }

                                PathNodeList.Add(n);
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotParse_Properly, "KMLPath" + ex.Message);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotParse_Properly", "KMLPath" + ex.Message);
                    return new List<ElementLayer>();
                }

                AllElementList = GetElementSurrondingEachPoint(ElementLayerList, PathNodeList);
            }

            return AllElementList.Distinct().ToList();
        }
        private List<Coord> ParseKMLPathCoord(string KMLTextPathForVector, List<ElementLayer> ElementLayerList)
        {
            string NotUsed = "";

            List<Coord> coordList = new List<Coord>();

            if (KMLTextPathForVector.Trim() == "")
            {
                return new List<Coord>();
            }
            else
            {
                try
                {
                    XmlReader reader = XmlReader.Create(new StringReader(KMLTextPathForVector));
                    while (reader.Read())
                    {
                        if (reader.Name == "coordinates")
                        {
                            string AllCoordinates = reader.ReadElementContentAsString().Trim();

                            string[] xyzArray = AllCoordinates.Split(" ".ToCharArray()[0]);
                            foreach (string xyz in xyzArray)
                            {
                                string[] xyzStr = xyz.Split(",".ToCharArray()[0]);
                                if (xyzStr.Count() != 3)
                                {
                                    return null;
                                }
                                Coord coord = new Coord();
                                if (Thread.CurrentThread.CurrentCulture.Name == "fr-CA")
                                {
                                    coord.Lng = float.Parse(xyzStr[0].Replace(".", ","));
                                    coord.Lat = float.Parse(xyzStr[1].Replace(".", ","));
                                    coord.Ordinal = 0;
                                }
                                else
                                {
                                    coord.Lng = float.Parse(xyzStr[0]);
                                    coord.Lat = float.Parse(xyzStr[1]);
                                    coord.Ordinal = 0;
                                }

                                coordList.Add(coord);
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    NotUsed = string.Format(TaskRunnerServiceRes.CouldNotParse_Properly, "KMLPath" + ex.Message);
                    _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("CouldNotParse_Properly", "KMLPath" + ex.Message);
                    return new List<Coord>();
                }
            }

            return coordList;
        }
        private bool PointInPolygon(Point p, Point[] poly)
        {
            Point p1, p2;
            bool inside = false;
            if (poly.Length < 3)
            {
                return inside;
            }

            Point oldPoint = new Point(poly[poly.Length - 1].X, poly[poly.Length - 1].Y);

            for (int i = 0; i < poly.Length; i++)
            {
                Point newPoint = new Point(poly[i].X, poly[i].Y);

                if (newPoint.X > oldPoint.X)
                {
                    p1 = oldPoint;
                    p2 = newPoint;
                }
                else
                {
                    p1 = newPoint;
                    p2 = oldPoint;
                }

                if ((newPoint.X < p.X) == (p.X <= oldPoint.X) && ((long)p.Y - (long)p1.Y) * (long)(p2.X - p1.X) < ((long)p2.Y - (long)p1.Y) * (long)(p.X - p1.X))
                {
                    inside = !inside;
                }

                oldPoint = newPoint;

            }
            return inside;
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
        private bool WriteKMLCurrentsAnimation(StringBuilder sbHTML, DfsuFile dfsuFile, List<ElementLayer> elementLayerList,
            List<NodeLayer> topNodeLayerList, List<NodeLayer> bottomNodeLayerList, List<ElementLayer> SelectedElementLayerList, string ProvInit)
        {
            string NotUsed = "";

            int ItemUVelocity = 0;
            int ItemVVelocity = 0;

            // getting the ItemNumber
            foreach (IDfsSimpleDynamicItemInfo dfsDyInfo in dfsuFile.ItemInfo)
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

            //int pcount = 0;
            sbHTML.AppendLine(@"<Folder><name>" + TaskRunnerServiceRes.CurrentsAnim + @"</name>");
            sbHTML.AppendLine(@"<visibility>0</visibility>");

            int CountLayer = (dfsuFile.NumberOfSigmaLayers == 0 ? 1 : dfsuFile.NumberOfSigmaLayers);


            for (int Layer = 1; Layer <= CountLayer; Layer++)
            {
                sbHTML.AppendLine(string.Format(@"<Folder><name>" + TaskRunnerServiceRes.Layer + " {0}</name>", Layer));
                sbHTML.AppendLine(@"<visibility>0</visibility>");

                Dictionary<int, Node> ElemCenter = new Dictionary<int, Node>();

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
                for (int timeStep = 0; timeStep < dfsuFile.NumberOfTimeSteps; timeStep++)
                {
                    sbHTML.AppendLine(@"<Folder>");
                    sbHTML.AppendLine(string.Format(@"<name>{0:yyyy-MM-dd} {0:HH:mm:ss tt}</name>", dfsuFile.StartDateTime.AddSeconds(vCount * dfsuFile.TimeStepInSeconds)));
                    sbHTML.AppendLine(@"<visibility>0</visibility>");
                    sbHTML.AppendLine(@"<TimeSpan>");

                    string Date_UTC_TextStart = "";
                    string Date_UTC_TextEnd = "";
                    if (ProvInit == "NL")
                    {
                        TimeZoneInfo tst = TimeZoneInfo.FindSystemTimeZoneById("Newfoundland Standard Time");
                        if (tst.IsDaylightSavingTime(dfsuFile.StartDateTime))
                        {
                            Date_UTC_TextStart = dfsuFile.StartDateTime.AddSeconds(vCount * dfsuFile.TimeStepInSeconds).AddHours(3).AddMinutes(30).ToString("yyyy-MM-ddTHH:mm:ss");
                            Date_UTC_TextEnd = dfsuFile.StartDateTime.AddSeconds((vCount + 1) * dfsuFile.TimeStepInSeconds).AddHours(3).AddMinutes(30).ToString("yyyy-MM-ddTHH:mm:ss");
                        }
                        else
                        {
                            Date_UTC_TextStart = dfsuFile.StartDateTime.AddSeconds(vCount * dfsuFile.TimeStepInSeconds).AddHours(2).AddMinutes(30).ToString("yyyy-MM-ddTHH:mm:ss");
                            Date_UTC_TextEnd = dfsuFile.StartDateTime.AddSeconds((vCount + 1) * dfsuFile.TimeStepInSeconds).AddHours(2).AddMinutes(30).ToString("yyyy-MM-ddTHH:mm:ss");
                        }
                    }
                    else if (ProvInit == "QC")
                    {
                        TimeZoneInfo tst = TimeZoneInfo.FindSystemTimeZoneById("Eastern Standard Time");
                        if (tst.IsDaylightSavingTime(dfsuFile.StartDateTime))
                        {
                            Date_UTC_TextStart = dfsuFile.StartDateTime.AddSeconds(vCount * dfsuFile.TimeStepInSeconds).AddHours(5).ToString("yyyy-MM-ddTHH:mm:ss");
                            Date_UTC_TextEnd = dfsuFile.StartDateTime.AddSeconds((vCount + 1) * dfsuFile.TimeStepInSeconds).AddHours(5).ToString("yyyy-MM-ddTHH:mm:ss");
                        }
                        else
                        {
                            Date_UTC_TextStart = dfsuFile.StartDateTime.AddSeconds(vCount * dfsuFile.TimeStepInSeconds).AddHours(4).ToString("yyyy-MM-ddTHH:mm:ss");
                            Date_UTC_TextEnd = dfsuFile.StartDateTime.AddSeconds((vCount + 1) * dfsuFile.TimeStepInSeconds).AddHours(4).ToString("yyyy-MM-ddTHH:mm:ss");
                        }
                    }
                    else if (ProvInit == "BC")
                    {
                        TimeZoneInfo tst = TimeZoneInfo.FindSystemTimeZoneById("Pacific Standard Time");
                        if (tst.IsDaylightSavingTime(dfsuFile.StartDateTime))
                        {
                            Date_UTC_TextStart = dfsuFile.StartDateTime.AddSeconds(vCount * dfsuFile.TimeStepInSeconds).AddHours(8).ToString("yyyy-MM-ddTHH:mm:ss");
                            Date_UTC_TextEnd = dfsuFile.StartDateTime.AddSeconds((vCount + 1) * dfsuFile.TimeStepInSeconds).AddHours(8).ToString("yyyy-MM-ddTHH:mm:ss");
                        }
                        else
                        {
                            Date_UTC_TextStart = dfsuFile.StartDateTime.AddSeconds(vCount * dfsuFile.TimeStepInSeconds).AddHours(7).ToString("yyyy-MM-ddTHH:mm:ss");
                            Date_UTC_TextEnd = dfsuFile.StartDateTime.AddSeconds((vCount + 1) * dfsuFile.TimeStepInSeconds).AddHours(7).ToString("yyyy-MM-ddTHH:mm:ss");
                        }
                    }
                    else
                    {
                        TimeZoneInfo tst = TimeZoneInfo.FindSystemTimeZoneById("Atlantic Standard Time");
                        if (tst.IsDaylightSavingTime(dfsuFile.StartDateTime))
                        {
                            Date_UTC_TextStart = dfsuFile.StartDateTime.AddSeconds(vCount * dfsuFile.TimeStepInSeconds).AddHours(4).ToString("yyyy-MM-ddTHH:mm:ss");
                            Date_UTC_TextEnd = dfsuFile.StartDateTime.AddSeconds((vCount + 1) * dfsuFile.TimeStepInSeconds).AddHours(4).ToString("yyyy-MM-ddTHH:mm:ss");
                        }
                        else
                        {
                            Date_UTC_TextStart = dfsuFile.StartDateTime.AddSeconds(vCount * dfsuFile.TimeStepInSeconds).AddHours(3).ToString("yyyy-MM-ddTHH:mm:ss");
                            Date_UTC_TextEnd = dfsuFile.StartDateTime.AddSeconds((vCount + 1) * dfsuFile.TimeStepInSeconds).AddHours(3).ToString("yyyy-MM-ddTHH:mm:ss");
                        }
                    }

                    sb.AppendLine($@"    <begin>{Date_UTC_TextStart}</begin>");
                    sb.AppendLine($@"    <end>{Date_UTC_TextEnd}</end>");

                    //sbHTML.AppendLine(string.Format(@"<begin>{0:yyyy-MM-dd}T{0:HH:mm:ss}</begin>", dfsuFile.StartDateTime.AddSeconds(vCount * dfsuFile.TimeStepInSeconds)));
                    //sbHTML.AppendLine(string.Format(@"<end>{0:yyyy-MM-dd}T{0:HH:mm:ss}</end>", dfsuFile.StartDateTime.AddSeconds((vCount + 1) * dfsuFile.TimeStepInSeconds)));
                    sbHTML.AppendLine(@"</TimeSpan>");

                    float[] UvelocityList = (float[])dfsuFile.ReadItemTimeStep(ItemUVelocity, timeStep).Data;
                    float[] VvelocityList = (float[])dfsuFile.ReadItemTimeStep(ItemVVelocity, timeStep).Data;

                    MapInfoService mapInfoService = new MapInfoService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
                    foreach (ElementLayer el in SelectedElementLayerList.Where(c => c.Layer == Layer))
                    {
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

                        if (VectorVal > 0)
                        {
                            sbHTML.AppendLine(@"<Placemark>");
                            sbHTML.AppendLine(@"<visibility>0</visibility>");
                            sbHTML.AppendLine(string.Format(@"<name>{0:F4} m/s " + TaskRunnerServiceRes.AtLowerCase + " {1:F0}°</name>", VectorVal, VectorDirCartesian).ToString().Replace(",", "."));

                            if (Layer == 1)
                            {
                                sbHTML.AppendLine(@"<styleUrl>#pink</styleUrl>");
                            }
                            else if (Layer == 2)
                            {
                                sbHTML.AppendLine(@"<styleUrl>#yellow</styleUrl>");
                            }
                            else if (Layer == 3)
                            {
                                sbHTML.AppendLine(@"<styleUrl>#green</styleUrl>");
                            }
                            else
                            {
                                // nothing ... It will be white by default
                            }

                            sbHTML.AppendLine(@"<LineString>");
                            sbHTML.AppendLine(@"<tessellate>1</tessellate>");
                            sbHTML.AppendLine(@"<coordinates>");

                            sbHTML.Append(((Node)ElemCenter[el.Element.ID]).X.ToString().Replace(",", ".") + @"," + ((Node)ElemCenter[el.Element.ID]).Y.ToString().Replace(",", ".") + ",0 ");

                            Node node = new Node();
                            double Fact = 0.00012;
                            double VectorSizeInMeterForEach10cm_s = 100;

                            double HypothDist = (VectorVal * VectorSizeInMeterForEach10cm_s * Fact);
                            node.X = (float)(ElemCenter[el.Element.ID].X + (HypothDist * Math.Cos(VectorDirCartesian * mapInfoService.d2r)));
                            node.Y = (float)(ElemCenter[el.Element.ID].Y + (HypothDist * Math.Sin(VectorDirCartesian * mapInfoService.d2r)));

                            sbHTML.Append(node.X.ToString().Replace(",", ".") + @"," + node.Y.ToString().Replace(",", ".") + ",0 ");

                            Node node2 = new Node();

                            node2.X = (float)(node.X + (HypothDist * 0.1 * Math.Cos((VectorDirCartesian + 180 - 25) * mapInfoService.d2r)));
                            node2.Y = (float)(node.Y + (HypothDist * 0.1 * Math.Sin((VectorDirCartesian + 180 - 25) * mapInfoService.d2r)));

                            sbHTML.Append(node2.X.ToString().Replace(",", ".") + @"," + node2.Y.ToString().Replace(",", ".") + ",0 ");
                            sbHTML.Append(node.X.ToString().Replace(",", ".") + @"," + node.Y.ToString().Replace(",", ".") + ",0 ");

                            node2.X = (float)(node.X + (HypothDist * 0.1 * Math.Cos((VectorDirCartesian + 180 + 25) * mapInfoService.d2r)));
                            node2.Y = (float)(node.Y + (HypothDist * 0.1 * Math.Sin((VectorDirCartesian + 180 + 25) * mapInfoService.d2r)));
                            sbHTML.Append(node2.X.ToString().Replace(",", ".") + @"," + node2.Y.ToString().Replace(",", ".") + ",0 ");

                            sbHTML.AppendLine(@"</coordinates>");
                            sbHTML.AppendLine(@"</LineString>");
                            sbHTML.AppendLine(@"</Placemark>");
                        }

                    }
                    sbHTML.AppendLine(@"</Folder>");
                    vCount += 1;
                }
                sbHTML.AppendLine(@"</Folder>");

            }
            sbHTML.AppendLine(@"</Folder>");

            return true;
        }
        private bool WriteKMLEstimatedDroguePathsAnimation(StringBuilder sbHTML, DfsuFile dfsuFile, List<ElementLayer> elementLayerList,
            List<NodeLayer> topNodeLayerList, List<NodeLayer> bottomNodeLayerList, List<ElementLayer> SelectedElementLayerList, List<Coord> StartCoordList,
            int DoFirstXDroguePoints, List<int> DelaysList, List<int> LayersList, string ProvInit)
        {
            string NotUsed = "";

            int ItemUVelocity = 0;
            int ItemVVelocity = 0;
            MapInfoService mapInfoService = new MapInfoService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);

            // getting the ItemNumber
            foreach (IDfsSimpleDynamicItemInfo dfsDyInfo in dfsuFile.ItemInfo)
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

            //int pcount = 0;
            sbHTML.AppendLine(@"<Folder><name>" + TaskRunnerServiceRes.EstimatedDroguePathsAnim + @"</name>");
            sbHTML.AppendLine(@"<visibility>0</visibility>");

            int CountLayer = (dfsuFile.NumberOfSigmaLayers == 0 ? 1 : dfsuFile.NumberOfSigmaLayers);

            Coord StartCoord = new Coord();
            Coord CurrentCoord = new Coord();

            for (int i = LayersList.Count - 1; i > 0; i--)
            {
                if (CountLayer < LayersList[i])
                {
                    LayersList.Remove(LayersList[i]);
                }
            }

            if (LayersList.Count == 0)
            {
                LayersList.Add(1);
            }

            int CountDrogue = 0;
            foreach (ElementLayer elementLayer in SelectedElementLayerList)
            {
                CountDrogue += 1;
                if (CountDrogue > DoFirstXDroguePoints)
                {
                    continue;
                }

                ElementLayer currentElementLayer = new ElementLayer()
                {
                    Element = elementLayer.Element,
                    Layer = elementLayer.Layer,
                    ZMax = elementLayer.ZMax,
                    ZMin = elementLayer.ZMin,
                };

                float XCenter = 0.0f;
                float YCenter = 0.0f;

                foreach (Node n in currentElementLayer.Element.NodeList)
                {
                    XCenter += n.X;
                    YCenter += n.Y;
                }

                XCenter = XCenter / currentElementLayer.Element.NodeList.Count();
                YCenter = YCenter / currentElementLayer.Element.NodeList.Count();

                StartCoord = new Coord() { Lat = StartCoordList[CountDrogue - 1].Lat, Lng = StartCoordList[CountDrogue - 1].Lng, Ordinal = 0 };

                sbHTML.AppendLine(@"<Placemark>");
                sbHTML.AppendLine(@"<visibility>1</visibility>");
                sbHTML.AppendLine($@"<name>{ TaskRunnerServiceRes.EstimatedDrogue } {CountDrogue } { TaskRunnerServiceRes.StartPosition }</name>");
                switch (CountDrogue)
                {
                    case 1:
                        {
                            sbHTML.AppendLine(@"<styleUrl>#msn_placemark_circle_green</styleUrl>");
                        }
                        break;
                    case 2:
                        {
                            sbHTML.AppendLine(@"<styleUrl>#msn_placemark_circle_red</styleUrl>");
                        }
                        break;
                    case 3:
                        {
                            sbHTML.AppendLine(@"<styleUrl>#msn_placemark_circle_blue</styleUrl>");
                        }
                        break;
                    case 4:
                        {
                            sbHTML.AppendLine(@"<styleUrl>#msn_placemark_circle_purple</styleUrl>");
                        }
                        break;
                    case 5:
                        {
                            sbHTML.AppendLine(@"<styleUrl>#msn_placemark_square_green</styleUrl>");
                        }
                        break;
                    case 6:
                        {
                            sbHTML.AppendLine(@"<styleUrl>#msn_placemark_square_red</styleUrl>");
                        }
                        break;
                    case 7:
                        {
                            sbHTML.AppendLine(@"<styleUrl>#msn_placemark_square_blue</styleUrl>");
                        }
                        break;
                    case 8:
                        {
                            sbHTML.AppendLine(@"<styleUrl>#msn_placemark_square_purple</styleUrl>");
                        }
                        break;
                    default:
                        {
                            sbHTML.AppendLine(@"<styleUrl>#msn_placemark_shaded_dot_green</styleUrl>");
                        }
                        break;
                }


                sbHTML.AppendLine(@"<Point>");
                sbHTML.AppendLine(@"<coordinates>");
                sbHTML.Append(@"");
                sbHTML.Append($@"{ StartCoord.Lng },{ StartCoord.Lat },0 ");
                sbHTML.AppendLine(@"</coordinates>");
                sbHTML.AppendLine(@"</Point>");
                sbHTML.AppendLine(@"</Placemark>");

            }

            CountDrogue = 0;
            int CountDelays = DelaysList.Count;
            int CountLayers = LayersList.Count;
            int CountSteps = 0;
            int TotalDrogueCount = DoFirstXDroguePoints;
            int TotalSteps = DoFirstXDroguePoints * CountDelays * CountLayers;
            foreach (ElementLayer elementLayer in SelectedElementLayerList)
            {
                CountDrogue += 1;
                if (CountDrogue > DoFirstXDroguePoints)
                {
                    continue;
                }

                ElementLayer currentElementLayer = new ElementLayer()
                {
                    Element = elementLayer.Element,
                    Layer = elementLayer.Layer,
                    ZMax = elementLayer.ZMax,
                    ZMin = elementLayer.ZMin,
                };

                float XCenter = 0.0f;
                float YCenter = 0.0f;

                foreach (Node n in currentElementLayer.Element.NodeList)
                {
                    XCenter += n.X;
                    YCenter += n.Y;
                }

                XCenter = XCenter / currentElementLayer.Element.NodeList.Count();
                YCenter = YCenter / currentElementLayer.Element.NodeList.Count();

                StartCoord = new Coord() { Lat = StartCoordList[CountDrogue - 1].Lat, Lng = StartCoordList[CountDrogue - 1].Lng, Ordinal = 0 };
                CurrentCoord = new Coord() { Lat = StartCoordList[CountDrogue - 1].Lat, Lng = StartCoordList[CountDrogue - 1].Lng, Ordinal = 0 };

                sbHTML.AppendLine(@"<Folder><name>" + TaskRunnerServiceRes.EstimatedDrogue + " " + CountDrogue.ToString() + @"</name>");
                sbHTML.AppendLine(@"<visibility>0</visibility>");

                for (int delayIndex = 0; delayIndex < DelaysList.Count; delayIndex++) // hrs to delay
                {
                    int delay = DelaysList[delayIndex];

                    int NumberOfTimeStepsToDelay = (int)(3600.0D / dfsuFile.TimeStepInSeconds * delay);

                    List<Coord> coordList = new List<Coord>()
                    {
                        StartCoord
                    };

                    sbHTML.AppendLine(string.Format(@"<Folder><name>" + TaskRunnerServiceRes.EstimatedDrogue + " {0} Delay {1} hrs</name>", CountDrogue, delay));
                    sbHTML.AppendLine(@"<visibility>0</visibility>");

                    for (int LayerIndex = 0; LayerIndex < LayersList.Count; LayerIndex++)
                    {
                        int Layer = LayersList[LayerIndex];

                        _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, (100 * CountSteps / TotalSteps));
                        _TaskRunnerBaseService.SendStatusTextToDB(_TaskRunnerBaseService.GetTextLanguageFormat1List("Creating_", $"Drogue { CountDrogue } of { TotalDrogueCount } Delay { delayIndex + 1 } of { CountDelays } Layer { LayerIndex + 1 } of { CountLayer }"));

                        CountSteps += 1;

                        sbHTML.AppendLine(string.Format(@"<Folder><name>" + TaskRunnerServiceRes.Layer + " {0}</name>", Layer));
                        sbHTML.AppendLine(@"<visibility>0</visibility>");

                        StringBuilder sbHTML2 = new StringBuilder();

                        int vCount = 0;
                        for (int timeStep = 0; timeStep < dfsuFile.NumberOfTimeSteps; timeStep++)
                        {
                            if (timeStep < NumberOfTimeStepsToDelay)
                            {
                                continue;
                            }

                            if (coordList.Count == 0)
                            {
                                NotUsed = string.Format(TaskRunnerServiceRes._ShouldNotBe0, "coordList.Count");
                                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_ShouldNotBe0", "coordList.Count");
                                return false;
                            }

                            List<Node> PathNodeList = new List<Node>()
                            {
                                new Node() { X = coordList[coordList.Count - 1].Lng, Y = coordList[coordList.Count - 1].Lat }
                            };

                            currentElementLayer = GetElementSurrondingEachPoint(elementLayerList, PathNodeList).FirstOrDefault();

                            if (currentElementLayer == null)
                            {
                                NotUsed = string.Format(TaskRunnerServiceRes._ShouldNotBeNull, "currentElementLayer");
                                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_ShouldNotBeNull", "currentElementLayer");
                                return false;
                            }

                            float[] UvelocityList = (float[])dfsuFile.ReadItemTimeStep(ItemUVelocity, timeStep).Data;
                            float[] VvelocityList = (float[])dfsuFile.ReadItemTimeStep(ItemVVelocity, timeStep).Data;

                            double totalTimeInSeconds = 0.0D;
                            double MinimumDistance = 10.0D;
                            while (totalTimeInSeconds < dfsuFile.TimeStepInSeconds)
                            {
                                float UV = UvelocityList[currentElementLayer.Element.ID - 1];
                                float VV = VvelocityList[currentElementLayer.Element.ID - 1];

                                double distLat = (double)VV * dfsuFile.TimeStepInSeconds;
                                double distLng = (double)UV * dfsuFile.TimeStepInSeconds;

                                double dist = Math.Sqrt((distLat * distLat) + (distLng * distLng));

                                if (dist != 0.0D)
                                {
                                    if (dist > MinimumDistance)
                                    {
                                        double fact = MinimumDistance / dist;
                                        double time = dfsuFile.TimeStepInSeconds * fact;
                                        totalTimeInSeconds += time;

                                        distLat = (double)VV * time;
                                        distLng = (double)UV * time;

                                        dist = Math.Sqrt((distLat * distLat) + (distLng * distLng));
                                    }
                                    else
                                    {
                                        totalTimeInSeconds = dfsuFile.TimeStepInSeconds;
                                    }

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


                                    double Lat1 = coordList[coordList.Count - 1].Lat;
                                    double Lng1 = coordList[coordList.Count - 1].Lng;

                                    Coord coord2 = _MapInfoService.CalculateDestination(Lat1 * _MapInfoService.d2r, Lng1 * _MapInfoService.d2r, dist, VectorDir * _MapInfoService.d2r);

                                    List<Node> PathNodeList2 = new List<Node>() { new Node() { X = coord2.Lng, Y = coord2.Lat } };
                                    ElementLayer NewElementLayer = GetElementSurrondingEachPoint(elementLayerList, new List<Node>() { new Node() { X = coord2.Lng, Y = coord2.Lat } }).FirstOrDefault();

                                    if (NewElementLayer == null)
                                    {
                                        NewElementLayer = currentElementLayer;
                                        coordList.Add(new Coord() { Lat = (float)Lat1, Lng = (float)Lng1, Ordinal = 0 });
                                        totalTimeInSeconds = dfsuFile.TimeStepInSeconds;
                                    }
                                    else
                                    {
                                        coordList.Add(new Coord() { Lat = coord2.Lat, Lng = coord2.Lng, Ordinal = 0 });
                                    }

                                    //if (currentElementLayer.Element.ID != NewElementLayer.Element.ID)
                                    //{

                                    sbHTML2.AppendLine(@"<Folder>");
                                    sbHTML2.AppendLine(string.Format(@"<name>{0:yyyy-MM-dd} {0:HH:mm:ss tt}</name>", dfsuFile.StartDateTime.AddSeconds((vCount * dfsuFile.TimeStepInSeconds))));
                                    sbHTML2.AppendLine(@"<visibility>0</visibility>");
                                    sbHTML2.AppendLine(@"<TimeSpan>");

                                    string Date_UTC_TextStart = "";
                                    string Date_UTC_TextEnd = "";
                                    if (ProvInit == "NL")
                                    {
                                        TimeZoneInfo tst = TimeZoneInfo.FindSystemTimeZoneById("Newfoundland Standard Time");
                                        if (tst.IsDaylightSavingTime(dfsuFile.StartDateTime))
                                        {
                                            Date_UTC_TextStart = dfsuFile.StartDateTime.AddSeconds((vCount * dfsuFile.TimeStepInSeconds) + totalTimeInSeconds).AddHours(3).AddMinutes(30).ToString("yyyy-MM-ddTHH:mm:ss");
                                            Date_UTC_TextEnd = dfsuFile.StartDateTime.AddSeconds(((vCount + 1) * dfsuFile.TimeStepInSeconds) + totalTimeInSeconds).AddHours(3).AddMinutes(30).ToString("yyyy-MM-ddTHH:mm:ss");
                                        }
                                        else
                                        {
                                            Date_UTC_TextStart = dfsuFile.StartDateTime.AddSeconds((vCount * dfsuFile.TimeStepInSeconds) + totalTimeInSeconds).AddHours(2).AddMinutes(30).ToString("yyyy-MM-ddTHH:mm:ss");
                                            Date_UTC_TextEnd = dfsuFile.StartDateTime.AddSeconds(((vCount + 1) * dfsuFile.TimeStepInSeconds) + totalTimeInSeconds).AddHours(2).AddMinutes(30).ToString("yyyy-MM-ddTHH:mm:ss");
                                        }
                                    }
                                    else if (ProvInit == "QC")
                                    {
                                        TimeZoneInfo tst = TimeZoneInfo.FindSystemTimeZoneById("Eastern Standard Time");
                                        if (tst.IsDaylightSavingTime(dfsuFile.StartDateTime))
                                        {
                                            Date_UTC_TextStart = dfsuFile.StartDateTime.AddSeconds((vCount * dfsuFile.TimeStepInSeconds) + totalTimeInSeconds).AddHours(5).ToString("yyyy-MM-ddTHH:mm:ss");
                                            Date_UTC_TextEnd = dfsuFile.StartDateTime.AddSeconds(((vCount + 1) * dfsuFile.TimeStepInSeconds) + totalTimeInSeconds).AddHours(5).ToString("yyyy-MM-ddTHH:mm:ss");
                                        }
                                        else
                                        {
                                            Date_UTC_TextStart = dfsuFile.StartDateTime.AddSeconds((vCount * dfsuFile.TimeStepInSeconds) + totalTimeInSeconds).AddHours(4).ToString("yyyy-MM-ddTHH:mm:ss");
                                            Date_UTC_TextEnd = dfsuFile.StartDateTime.AddSeconds(((vCount + 1) * dfsuFile.TimeStepInSeconds) + totalTimeInSeconds).AddHours(4).ToString("yyyy-MM-ddTHH:mm:ss");
                                        }
                                    }
                                    else if (ProvInit == "BC")
                                    {
                                        TimeZoneInfo tst = TimeZoneInfo.FindSystemTimeZoneById("Pacific Standard Time");
                                        if (tst.IsDaylightSavingTime(dfsuFile.StartDateTime))
                                        {
                                            Date_UTC_TextStart = dfsuFile.StartDateTime.AddSeconds((vCount * dfsuFile.TimeStepInSeconds) + totalTimeInSeconds).AddHours(8).ToString("yyyy-MM-ddTHH:mm:ss");
                                            Date_UTC_TextEnd = dfsuFile.StartDateTime.AddSeconds(((vCount + 1) * dfsuFile.TimeStepInSeconds) + totalTimeInSeconds).AddHours(8).ToString("yyyy-MM-ddTHH:mm:ss");
                                        }
                                        else
                                        {
                                            Date_UTC_TextStart = dfsuFile.StartDateTime.AddSeconds((vCount * dfsuFile.TimeStepInSeconds) + totalTimeInSeconds).AddHours(7).ToString("yyyy-MM-ddTHH:mm:ss");
                                            Date_UTC_TextEnd = dfsuFile.StartDateTime.AddSeconds(((vCount + 1) * dfsuFile.TimeStepInSeconds) + totalTimeInSeconds).AddHours(7).ToString("yyyy-MM-ddTHH:mm:ss");
                                        }
                                    }
                                    else
                                    {
                                        TimeZoneInfo tst = TimeZoneInfo.FindSystemTimeZoneById("Atlantic Standard Time");
                                        if (tst.IsDaylightSavingTime(dfsuFile.StartDateTime))
                                        {
                                            Date_UTC_TextStart = dfsuFile.StartDateTime.AddSeconds((vCount * dfsuFile.TimeStepInSeconds) + totalTimeInSeconds).AddHours(4).ToString("yyyy-MM-ddTHH:mm:ss");
                                            Date_UTC_TextEnd = dfsuFile.StartDateTime.AddSeconds(((vCount + 1) * dfsuFile.TimeStepInSeconds) + totalTimeInSeconds).AddHours(4).ToString("yyyy-MM-ddTHH:mm:ss");
                                        }
                                        else
                                        {
                                            Date_UTC_TextStart = dfsuFile.StartDateTime.AddSeconds((vCount * dfsuFile.TimeStepInSeconds) + totalTimeInSeconds).AddHours(3).ToString("yyyy-MM-ddTHH:mm:ss");
                                            Date_UTC_TextEnd = dfsuFile.StartDateTime.AddSeconds(((vCount + 1) * dfsuFile.TimeStepInSeconds) + totalTimeInSeconds).AddHours(3).ToString("yyyy-MM-ddTHH:mm:ss");
                                        }
                                    }

                                    sbHTML2.AppendLine($@"    <begin>{Date_UTC_TextStart}</begin>");
                                    sbHTML2.AppendLine($@"    <end>{Date_UTC_TextEnd}</end>");

                                    //sbHTML2.AppendLine(string.Format(@"<begin>{0:yyyy-MM-dd}T{0:HH:mm:ss}</begin>", dfsuFile.StartDateTime.AddSeconds((vCount * dfsuFile.TimeStepInSeconds) + totalTimeInSeconds)));
                                    //sbHTML2.AppendLine(string.Format(@"<end>{0:yyyy-MM-dd}T{0:HH:mm:ss}</end>", dfsuFile.StartDateTime.AddSeconds(((vCount + 1) * dfsuFile.TimeStepInSeconds) + totalTimeInSeconds)));
                                    sbHTML2.AppendLine(@"</TimeSpan>");
                                    sbHTML2.AppendLine(@"<Placemark>");
                                    sbHTML2.AppendLine(@"<visibility>0</visibility>");
                                    sbHTML2.AppendLine($@"<name>EDP { CountDrogue }</name>");

                                    switch (CountDrogue)
                                    {
                                        case 1:
                                            {
                                                sbHTML2.AppendLine(@"<styleUrl>#msn_placemark_circle_green</styleUrl>");
                                            }
                                            break;
                                        case 2:
                                            {
                                                sbHTML2.AppendLine(@"<styleUrl>#msn_placemark_circle_red</styleUrl>");
                                            }
                                            break;
                                        case 3:
                                            {
                                                sbHTML2.AppendLine(@"<styleUrl>#msn_placemark_circle_blue</styleUrl>");
                                            }
                                            break;
                                        case 4:
                                            {
                                                sbHTML2.AppendLine(@"<styleUrl>#msn_placemark_circle_purple</styleUrl>");
                                            }
                                            break;
                                        case 5:
                                            {
                                                sbHTML2.AppendLine(@"<styleUrl>#msn_placemark_square_green</styleUrl>");
                                            }
                                            break;
                                        case 6:
                                            {
                                                sbHTML2.AppendLine(@"<styleUrl>#msn_placemark_square_red</styleUrl>");
                                            }
                                            break;
                                        case 7:
                                            {
                                                sbHTML2.AppendLine(@"<styleUrl>#msn_placemark_square_blue</styleUrl>");
                                            }
                                            break;
                                        case 8:
                                            {
                                                sbHTML2.AppendLine(@"<styleUrl>#msn_placemark_square_purple</styleUrl>");
                                            }
                                            break;
                                        default:
                                            {
                                                sbHTML2.AppendLine(@"<styleUrl>#msn_placemark_shaded_dot_green</styleUrl>");
                                            }
                                            break;
                                    }

                                    sbHTML2.AppendLine(@"<Point>");
                                    sbHTML2.AppendLine(@"<coordinates>");
                                    sbHTML2.Append(@"");
                                    if (coordList.Count > 1)
                                    {
                                        sbHTML2.Append($@"{ coordList[coordList.Count - 1].Lng },{ coordList[coordList.Count - 1].Lat },0 ");
                                    }
                                    sbHTML2.AppendLine(@"</coordinates>");
                                    sbHTML2.AppendLine(@"</Point>");
                                    sbHTML2.AppendLine(@"</Placemark>");
                                    sbHTML2.AppendLine(@"</Folder>");

                                    //}
                                }
                                else
                                {
                                    totalTimeInSeconds = dfsuFile.TimeStepInSeconds;
                                }
                            }
                            vCount += 1;
                        }

                        sbHTML.AppendLine(@"<Folder>");
                        sbHTML.AppendLine(@"<name>Full Path</name>");
                        sbHTML.AppendLine(@"<visibility>0</visibility>");
                        sbHTML.AppendLine(@"<Placemark>");
                        sbHTML.AppendLine(@"<visibility>0</visibility>");
                        sbHTML.AppendLine($@"<name>EDP { CountDrogue }</name>");
                        switch (CountDrogue)
                        {
                            case 1:
                                {
                                    sbHTML.AppendLine(@"<styleUrl>#msn_placemark_circle_green</styleUrl>");
                                }
                                break;
                            case 2:
                                {
                                    sbHTML.AppendLine(@"<styleUrl>#msn_placemark_circle_red</styleUrl>");
                                }
                                break;
                            case 3:
                                {
                                    sbHTML.AppendLine(@"<styleUrl>#msn_placemark_circle_blue</styleUrl>");
                                }
                                break;
                            case 4:
                                {
                                    sbHTML.AppendLine(@"<styleUrl>#msn_placemark_circle_purple</styleUrl>");
                                }
                                break;
                            case 5:
                                {
                                    sbHTML.AppendLine(@"<styleUrl>#msn_placemark_square_green</styleUrl>");
                                }
                                break;
                            case 6:
                                {
                                    sbHTML.AppendLine(@"<styleUrl>#msn_placemark_square_red</styleUrl>");
                                }
                                break;
                            case 7:
                                {
                                    sbHTML.AppendLine(@"<styleUrl>#msn_placemark_square_blue</styleUrl>");
                                }
                                break;
                            case 8:
                                {
                                    sbHTML.AppendLine(@"<styleUrl>#msn_placemark_square_purple</styleUrl>");
                                }
                                break;
                            default:
                                {
                                    sbHTML.AppendLine(@"<styleUrl>#msn_placemark_shaded_dot_green</styleUrl>");
                                }
                                break;
                        }
                        sbHTML.AppendLine(@"<LineString>");
                        sbHTML.AppendLine(@"<tessellate>1</tessellate>");
                        sbHTML.AppendLine(@"<coordinates>");
                        sbHTML.Append(@"");
                        foreach (Coord coord in coordList)
                        {
                            sbHTML.Append($@"{ coord.Lng },{ coord.Lat },0 ");

                        }
                        sbHTML.AppendLine(@"</coordinates>");
                        sbHTML.AppendLine(@"</LineString>");
                        sbHTML.AppendLine(@"</Placemark>");
                        sbHTML.AppendLine(@"</Folder>");

                        sbHTML.AppendLine(sbHTML2.ToString());

                        sbHTML.AppendLine(@"</Folder>"); // TaskRunnerServiceRes.Layer

                    }

                    sbHTML.AppendLine(@"</Folder>"); // Delay

                }

                sbHTML.AppendLine(@"</Folder>"); // TaskRunnerServiceRes.EstimatedDroguePathsAnim

            }

            sbHTML.AppendLine(@"</Folder>"); // TaskRunnerServiceRes.EstimatedDroguePathsAnim

            return true;
        }
        private bool WriteKMLFecalColiformContourLine(StringBuilder sbHTML, DfsuFile dfsuFile, List<ElementLayer> elementLayerList, List<NodeLayer> topNodeLayerList, List<NodeLayer> bottomNodeLayerList, List<float> ContourValueList, string ProvInit)
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
            sbHTML.AppendLine(@"<Folder>");
            sbHTML.AppendLine(@"  <name>" + TaskRunnerServiceRes.PollutionAnimation + "</name>");
            sbHTML.AppendLine(@"  <visibility>0</visibility>");

            int CountAt = 0;
            int CountLayer = (dfsuFile.NumberOfSigmaLayers == 0 ? 1 : dfsuFile.NumberOfSigmaLayers);
            int CurrentContourValue = 0;
            //int CurrentTimeSteps = 0;

            int TotalCount = CountLayer * ContourValueList.Count * dfsuFile.NumberOfTimeSteps;

            for (int Layer = 1; Layer <= CountLayer; Layer++)
            {
                //CurrentLayer += 1;
                CurrentContourValue = 1;
                //CurrentTimeSteps = 1;

                #region Top of Layer
                if (Layer == 1)
                {
                    sbHTML.AppendLine(string.Format(@"<Folder><name>" + TaskRunnerServiceRes.TopOfLayer + @" [{0}] </name>", Layer));
                }
                else
                {
                    sbHTML.AppendLine(string.Format(@"<Folder><name>" + TaskRunnerServiceRes.TopOfLayer + @" [{0}] " + TaskRunnerServiceRes.BottomOfLayer + @" [{1}] </name>", Layer, Layer - 1));
                }
                sbHTML.AppendLine(@"<visibility>0</visibility>");
                int CountContourValue = 1;
                foreach (float ContourValue in ContourValueList)
                {
                    sbHTML.AppendLine(string.Format(@"  <Folder><name>" + TaskRunnerServiceRes.ContourValue + @" [{0}]</name>", ContourValue));
                    sbHTML.AppendLine(@"  <visibility>0</visibility>");

                    int vcount = 0;
                    CurrentContourValue += 1;
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
                            if (dfsuFile.DfsuFileType == DfsuFileType.Dfsu3DSigmaZ)
                            {
                                NotUsed = string.Format(TaskRunnerServiceRes._NotImplemented, "Z Level");
                                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_NotImplemented", "Z Level");
                                return false;
                            }

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

                        UniqueElementList = (from el in TempUniqueElementList select el).Distinct().ToList<Element>();

                        // filling InterpolatedContourNodeList
                        InterpolatedContourNodeList = new List<Node>();

                        foreach (Element el in UniqueElementList)
                        {
                            if (dfsuFile.DfsuFileType == DfsuFileType.Dfsu3DSigmaZ)
                            {
                                NotUsed = string.Format(TaskRunnerServiceRes._NotImplemented, "Z Level");
                                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_NotImplemented", "Z Level");
                                return false;
                            }

                            if (dfsuFile.DfsuFileType == DfsuFileType.Dfsu3DSigma)
                            {
                                if (el.Type == 32)
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
                                else if (el.Type == 33)
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
                                FillVectors21_32(el, UniqueElementList, ContourValue);
                            }
                            else if (el.Type == 24)
                            {
                                NotUsed = TaskRunnerServiceRes.AllNodesAreSmallerThanContourValue;
                                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageList("AllNodesAreSmallerThanContourValue");
                                return false;
                            }
                            else if (el.Type == 25)
                            {
                                FillVectors25_33(el, UniqueElementList, ContourValue);
                            }
                            else if (el.Type == 32)
                            {
                                FillVectors21_32(el, UniqueElementList, ContourValue);
                            }
                            else if (el.Type == 33)
                            {
                                FillVectors25_33(el, UniqueElementList, ContourValue);
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
                        DrawKMLContourPolygon(ContourPolygonList, dfsuFile, vcount, ProvInit);

                        vcount += 1;
                    }
                    sbHTML.AppendLine(@"  </Folder>");
                    CountContourValue += 1;
                }
                sbHTML.AppendLine(@"  </Folder>");
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
            sbHTML.AppendLine(@"</Folder>");

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

            for (int Layer = 1; Layer <= CountLayer; Layer++)
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
                        if (dfsuFile.DfsuFileType == DfsuFileType.Dfsu3DSigmaZ)
                        {
                            NotUsed = string.Format(TaskRunnerServiceRes._NotImplemented, "Z Level");
                            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_NotImplemented", "Z Level");
                            return false;
                        }

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

                    UniqueElementList = (from el in TempUniqueElementList select el).Distinct().ToList<Element>();

                    // filling InterpolatedContourNodeList
                    InterpolatedContourNodeList = new List<Node>();

                    int count = 0;
                    foreach (Element el in UniqueElementList)
                    {
                        if (dfsuFile.DfsuFileType == DfsuFileType.Dfsu3DSigmaZ)
                        {
                            NotUsed = string.Format(TaskRunnerServiceRes._NotImplemented, "Z Level");
                            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("_NotImplemented", "Z Level");
                            return false;
                        }

                        count += 1;
                        if (dfsuFile.DfsuFileType == DfsuFileType.Dfsu3DSigma)
                        {
                            if (el.Type == 32)
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
                            else if (el.Type == 33)
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
                            FillVectors21_32(el, UniqueElementList, ContourValue);
                        }
                        else if (el.Type == 24)
                        {
                            NotUsed = string.Format(TaskRunnerServiceRes.ElementType_IsNotSupported, el.Type.ToString());
                            _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat1List("ElementType_IsNotSupported", el.Type.ToString());
                            return false;
                        }
                        else if (el.Type == 25)
                        {
                            FillVectors25_33(el, UniqueElementList, ContourValue);
                        }
                        else if (el.Type == 32)
                        {
                            FillVectors21_32(el, UniqueElementList, ContourValue);
                        }
                        else if (el.Type == 33)
                        {
                            FillVectors25_33(el, UniqueElementList, ContourValue);
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

            foreach (ContourPolygon contourPolygon in ContourPolygonList)
            {
                bool Found = true;
                while (Found)
                {
                    if (contourPolygon.ContourNodeList[0].Code != 1)
                    {
                        Node n = contourPolygon.ContourNodeList[0];
                        contourPolygon.ContourNodeList.RemoveAt(0);
                        contourPolygon.ContourNodeList.Add(contourPolygon.ContourNodeList[0]);
                    }
                    else
                    {
                        Found = false;
                    }
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
