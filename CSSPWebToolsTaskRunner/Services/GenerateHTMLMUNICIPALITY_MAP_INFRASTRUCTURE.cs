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
using CSSPEnumsDLL.Services;
using CSSPModelsDLL.Models;
using System.Windows.Forms;
//using System.Web.Helpers;

namespace CSSPWebToolsTaskRunner.Services
{
    public partial class ParametersService
    {
        private bool GenerateHTMLMUNICIPALITY_MAP_INFRASTRUCTURE(StringBuilder sbTemp)
        {
            int Percent = 10;
            string NotUsed = "";

            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, Percent);
            _TaskRunnerBaseService.SendStatusTextToDB(_TaskRunnerBaseService.GetTextLanguageFormat1List("Creating_", ReportGenerateObjectsKeywordEnum.SUBSECTOR_MWQM_SITES_FC_TABLE.ToString()));

            List<string> ParamValueList = Parameters.Split("|||".ToCharArray(), StringSplitOptions.RemoveEmptyEntries).ToList();

            // TVItemID and Year already loaded

            TVItemModel tvItemModelSubsector = _TVItemService.GetTVItemModelWithTVItemIDDB(TVItemID);
            if (!string.IsNullOrWhiteSpace(tvItemModelSubsector.Error))
            {
                NotUsed = string.Format(TaskRunnerServiceRes.CouldNotFind_With_Equal_, TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, TVItemID.ToString());
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageFormat3List("CouldNotFind_With_Equal_", TaskRunnerServiceRes.TVItem, TaskRunnerServiceRes.TVItemID, TVItemID.ToString());
                return false;
            }

            string ServerPath = _TVFileService.GetServerFilePath(tvItemModelSubsector.TVItemID);

            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 5);

            List<TVItemModel> tvItemModelListMunicipality = _TVItemService.GetChildrenTVItemModelListWithTVItemIDAndTVTypeDB(tvItemModelSubsector.TVItemID, TVTypeEnum.Municipality).Where(c => c.IsActive == true).ToList();

            if (tvItemModelListMunicipality.Count > 0)
            {
                foreach (TVItemModel tvItemModel in tvItemModelListMunicipality)
                {
                    sbTemp.AppendLine($@"<h3>{ tvItemModel.TVText } map of infrastructure</h3");

                    sbTemp.AppendLine($@"<p>{ TaskRunnerServiceRes.NotImplementedYet }</p>");

                }
            }
            else
            {
                sbTemp.AppendLine($@"<p style=""font-size: 2em;"">{ TaskRunnerServiceRes.NoMunicipalityWithinSubsector }</p>");
            }

            _TaskRunnerBaseService.SendPercentToDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID, 80);

            return true;
        }

        // for testing only can comment out when test is completed
        public bool PublicGenerateHTMLMUNICIPALITY_MAP_INFRASTRUCTURE(StringBuilder sbTemp)
        {
            bool retBool = GenerateHTMLMUNICIPALITY_MAP_INFRASTRUCTURE(sbTemp);

            StreamWriter sw = fi.CreateText();
            sw.Write(sbTemp.ToString());
            sw.Flush();
            sw.Close();

            return retBool;
        }
    }
}
