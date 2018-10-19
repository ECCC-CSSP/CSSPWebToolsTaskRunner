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

namespace CSSPWebToolsTaskRunner.Services
{
    public partial class ParametersService
    {
        private bool GenerateKMZMWQMRun(FileInfo fi, StringBuilder sbKMZ, string parameters, ReportTypeModel reportTypeModel)
        {
            switch (reportTypeModel.UniqueCode)
            {
                case "MWQMRunTestKMZ":
                    {
                        if (!GenerateKMZMWQMRun_MWQMRunTestKMZ(fi, sbKMZ, parameters, reportTypeModel))
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
        private bool GenerateKMZMWQMRun_MWQMRunTestKMZ(FileInfo fi, StringBuilder sbKMZ, string parameters, ReportTypeModel reportTypeModel)
        {
            sbKMZ.AppendLine(@"<?xml version=""1.0"" encoding=""UTF-8""?>");
            sbKMZ.AppendLine(@"<kml xmlns=""http://www.opengis.net/kml/2.2"" xmlns:gx=""http://www.google.com/kml/ext/2.2"" xmlns:kml=""http://www.opengis.net/kml/2.2"" xmlns:atom=""http://www.w3.org/2005/Atom"">");
            sbKMZ.AppendLine(@"<Document>");
            sbKMZ.AppendLine(@"	<name>" + fi.FullName.Replace(".kml", ".kmz") + "</name>");
            sbKMZ.AppendLine(@"	<Placemark>");
            sbKMZ.AppendLine(@"		<name>My Parents Home</name> ");
            sbKMZ.AppendLine(@"		<Point>");
            sbKMZ.AppendLine(@"			<coordinates>-64.69002452357093,46.48465663502946,0</coordinates>");
            sbKMZ.AppendLine(@"		</Point> ");
            sbKMZ.AppendLine(@"	</Placemark>");
            sbKMZ.AppendLine(@"</Document> ");
            sbKMZ.AppendLine(@"</kml>");

            return true;
        }
        private bool GenerateKMZMWQMRun_NotImplementedKMZ(FileInfo fi, StringBuilder sbKMZ, string parameters, ReportTypeModel reportTypeModel)
        {
            sbKMZ.AppendLine(@"<?xml version=""1.0"" encoding=""UTF-8""?>");
            sbKMZ.AppendLine(@"<kml xmlns=""http://www.opengis.net/kml/2.2"" xmlns:gx=""http://www.google.com/kml/ext/2.2"" xmlns:kml=""http://www.opengis.net/kml/2.2"" xmlns:atom=""http://www.w3.org/2005/Atom"">");
            sbKMZ.AppendLine(@"<Document>");
            sbKMZ.AppendLine(@"	<name>" + fi.FullName.Replace(".kml", ".kmz") + "</name>");
            sbKMZ.AppendLine(@"	<Placemark>");
            sbKMZ.AppendLine(@"		<name>Not Implemented</name> ");
            sbKMZ.AppendLine(@"		<Point>");
            sbKMZ.AppendLine(@"			<coordinates>-90,50,0</coordinates>");
            sbKMZ.AppendLine(@"		</Point> ");
            sbKMZ.AppendLine(@"	</Placemark>");
            sbKMZ.AppendLine(@"</Document> ");
            sbKMZ.AppendLine(@"</kml>");

            return true;
        }
    }
}
