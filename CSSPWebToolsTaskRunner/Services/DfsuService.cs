using DHI.Generic.MikeZero.DFS.dfsu;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CSSPWebToolsTaskRunner.Services;
using CSSPWebToolsTaskRunner.Services.Resources;
using System.Windows.Forms;
using System.Drawing;
using System.Diagnostics;
using System.Transactions;
using System.Xml;
using System.Threading;
using DHI.Generic.MikeZero.DFS;
using DHI.Generic.MikeZero;
using CSSPWebToolsDBDLL.Models;
using CSSPWebToolsDBDLL;
using CSSPWebToolsDBDLL.Services;
using CSSPWebToolsDBDLL.Services.Resources;
using System.Globalization;
using CSSPModelsDLL.Models;

namespace CSSPWebToolsTaskRunner.Services
{
    public class DfsuService
    {
        #region Variables
        public TaskRunnerBaseService _TaskRunnerBaseService { get; private set; }
        public DfsuFile _DfsuFile;
        public string _DfsuFileName { get; private set; }

        public List<Node> InterpolatedContourNodeList { get; set; }
        public Dictionary<String, Vector> ForwardVector { get; set; }
        public Dictionary<String, Vector> BackwardVector { get; set; }
        public List<Element> ElementList { get; set; }
        #endregion Variables

        #region Constructors
        public DfsuService(TaskRunnerBaseService taskRunnerBaseService, string DfsuFileName)
        {
            _TaskRunnerBaseService = taskRunnerBaseService;
            _DfsuFileName = DfsuFileName;
            InterpolatedContourNodeList = new List<Node>();
            ForwardVector = new Dictionary<string, Vector>();
            BackwardVector = new Dictionary<string, Vector>();
            ElementList = new List<Element>();
        }
        #endregion Constructors

        #region Functions public

        #endregion Functions public

        #region Functions private
        #endregion Functions private

    }
}
