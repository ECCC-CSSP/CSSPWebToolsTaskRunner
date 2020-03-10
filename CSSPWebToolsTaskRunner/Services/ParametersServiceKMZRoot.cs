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
        private bool GenerateKMZRoot()
        {
            switch (reportTypeModel.UniqueCode)
            {
                case "RootTestFRKMZ":
                case "RootTestENKMZ":
                    {
                        if (!GenerateKMZRoot_RootTestKMZ())
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
        private bool GenerateKMZRoot_RootTestKMZ()
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
        private bool GenerateKMZRoot_NotImplementedKMZ()
        {
            sb.AppendLine(@"<?xml version=""1.0"" encoding=""UTF-8""?>");
            sb.AppendLine(@"<kml xmlns=""http://www.opengis.net/kml/2.2"" xmlns:gx=""http://www.google.com/kml/ext/2.2"" xmlns:kml=""http://www.opengis.net/kml/2.2"" xmlns:atom=""http://www.w3.org/2005/Atom"">");
            sb.AppendLine(@"<Document>");
            sb.AppendLine(@"	<name>" + fi.FullName.Replace(".kml", ".kmz") + "</name>");
            sb.AppendLine(@"	<Placemark>");
            sb.AppendLine(@"		<name>Not Implemented</name> ");
            sb.AppendLine(@"		<Point>");
            sb.AppendLine(@"			<coordinates>-90,50,0</coordinates>");
            sb.AppendLine(@"		</Point> ");
            sb.AppendLine(@"	</Placemark>");
            sb.AppendLine(@"</Document> ");
            sb.AppendLine(@"</kml>");

            return true;
        }
    }
}
