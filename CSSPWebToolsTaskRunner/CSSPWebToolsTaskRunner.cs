using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using CSSPWebToolsDBDLL.Models;
using CSSPWebToolsDBDLL.Services;
using CSSPWebToolsTaskRunner.Services;
using System.Security.Principal;
using CSSPWebToolsTaskRunner.Services.Resources;
using System.Threading;
using System.Globalization;
using System.Net;
using System.Collections.Specialized;
using System.Net.Mail;
using CSSPWebToolsDBDLL;
using System.IO;
using System.Transactions;
using CSSPModelsDLL.Models;
using CSSPEnumsDLL.Enums;

namespace CSSPWebToolsTaskRunner
{
    public partial class CSSPWebToolsTaskRunner : Form
    {
        #region Variables
        private int LabSheetLookupDelay = 3; // seconds
        private int TaskStatusOfRunnningLookupDelay = 5; // seconds
        private List<int> DavidBenoitEmailTimeHourList = new List<int>() { 6, 12, 18 }; // hours to send email every day
        private int MPNLimitForEmail = 500;
        private bool testing = false;
        #endregion Variables

        #region Properties
        public int _BWCount { get; set; }
        public List<BWObj> _BWList { get; set; }
        public RichTextBox _RichTextBoxStatus { get; set; }
        public int _SkipTimerCount { get; set; }
        public TaskRunnerBaseService _TaskRunnerBaseService { get; set; }
        public System.Windows.Forms.Timer _TimerCheckTask { get; set; }
        public IPrincipal _User { get; set; }
        public Label _LabelLastAppTaskCheckDate { get; set; }
        #endregion Properties

        #region Constructors
        public CSSPWebToolsTaskRunner()
        {
            InitializeComponent();
            AppInit();
            _User = new GenericPrincipal(new GenericIdentity("Charles.LeBlanc2@Canada.ca", "Forms"), null);
        }
        #endregion Constructors

        #region Events
        private void bw_DoWork(object sender, DoWorkEventArgs e)
        {
            BWDoWork((BWObj)e.Argument);
        }
        private void timerCheckTask_Tick(object sender, EventArgs e)
        {
            DoTimerCheckTask();
        }
        #endregion Events

        #region Functions
        public void AppInit()
        {
            _BWCount = 100;
            _SkipTimerCount = 0;
            _TimerCheckTask = timerCheckTask;
            richTextBoxStatus.Text = "";
            _BWList = new List<BWObj>();
            for (int i = 0; i < _BWCount; i++)
            {
                _BWList.Add(new BWObj());
            }
            foreach (BWObj bwObj in _BWList)
            {
                bwObj.bw.DoWork += bw_DoWork;
            }
            _TaskRunnerBaseService = new TaskRunnerBaseService(_BWList);
            _TaskRunnerBaseService._RichTextBoxStatus = richTextBoxStatus;
            timerCheckTask.Enabled = true;

            _RichTextBoxStatus = richTextBoxStatus;
            _LabelLastAppTaskCheckDate = lblLastAppTaskCheckDate;
        }
        public void BWDoWork(BWObj bwObj)
        {
            string NotUsed = "";

            if (Enum.GetNames(typeof(AppTaskCommandEnum)).Contains(bwObj.appTaskCommand.ToString()))
            {
                _TaskRunnerBaseService.ExecuteTask();
                //_TaskRunnerBaseService.GenerateDoc();
                //_TaskRunnerBaseService.DoCommand();
            }
            else
            {
                AppTaskService appTaskService = new AppTaskService(_TaskRunnerBaseService._BWObj.appTaskModel.Language, _TaskRunnerBaseService._User);
                AppTaskModel appTaskModel = appTaskService.GetAppTaskModelWithAppTaskIDDB(_TaskRunnerBaseService._BWObj.appTaskModel.AppTaskID);
                NotUsed = TaskRunnerServiceRes.GenerateOrCommandNeedsToBeTrue;
                _TaskRunnerBaseService._BWObj.TextLanguageList = _TaskRunnerBaseService.GetTextLanguageList(TaskRunnerServiceRes.GenerateOrCommandNeedsToBeTrue);
                if (_TaskRunnerBaseService._BWObj.TextLanguageList.Count == 0)
                {
                    appTaskService.PostDeleteAppTaskDB(appTaskModel.AppTaskID);
                }
                else
                {
                    _TaskRunnerBaseService.SendErrorTextToDB(_TaskRunnerBaseService._BWObj.TextLanguageList);
                }
            }
        }
        public void BWTryRunningTask(AppTaskModel appTaskModel)
        {
            bool BWAvailable = false;
            int Ind = 0;

            AppTaskService appTaskService = new AppTaskService(appTaskModel.Language, _User);

            foreach (BWObj bwObj in _BWList)
            {
                Ind += 1;
                if (!bwObj.bw.IsBusy)
                {
                    bwObj.Index = Ind;
                    bwObj.appTaskModel = appTaskModel;
                    bwObj.appTaskCommand = appTaskModel.AppTaskCommand;
                    _TaskRunnerBaseService._BWObj = bwObj;
                    bwObj.TextLanguageList = new List<TextLanguage>();
                    bwObj.bw.RunWorkerAsync(bwObj);
                    BWAvailable = true;

                    WriteToRichTextBoxStatus(bwObj.appTaskCommand + string.Format(TaskRunnerServiceRes.StartingWithBackgroundWorkingIndex_, bwObj.Index) + "\r\n");
                    break;
                }
            }

            if (!BWAvailable)
            {
                WriteToRichTextBoxStatus(string.Format(TaskRunnerServiceRes.CouldNotFindAvailableBackgroundWorkerToRunAppTaskID_AppTaskCommand_, appTaskModel.AppTaskID, appTaskModel.AppTaskCommand.ToString()) + "\r\n");
            }

            return;
        }
        public void DoTimerCheckTask()
        {
            timerCheckTask.Enabled = false;
            GetNextTask();
            timerCheckTask.Enabled = true;
        }
        public void GetNextTask()
        {
            if (_SkipTimerCount == 86401)
            {
                _SkipTimerCount = 0;
            }
            _SkipTimerCount += 1;

            lblLastAppTaskCheckDate.Text = DateTime.Now.ToString("F");

            AppTaskService appTaskService = new AppTaskService(LanguageEnum.en, _User);

            AppTaskModel appTaskModel = appTaskService.CheckAppTask();

            if (!string.IsNullOrWhiteSpace(appTaskModel.Error))
            {
                lblLastAppTaskCheckDate.Text += " " + appTaskModel.Error;
            }
            else
            {
                appTaskModel.AppTaskStatus = AppTaskStatusEnum.Running;
                appTaskModel = appTaskService.PostUpdateAppTask(appTaskModel);
                if (!string.IsNullOrWhiteSpace(appTaskModel.Error))
                    return;

                BWTryRunningTask(appTaskModel);
            }

            // Checking for new lab sheets
            if (_SkipTimerCount % LabSheetLookupDelay == 0)
            {
                GetNextLabSheet();
            }

            // Checking others
            if (_SkipTimerCount % TaskStatusOfRunnningLookupDelay == 0) // if timer interval is 1 second then this will run every 5 seconds
            {
                _TaskRunnerBaseService.UpdateStatusOfRunningScenarios();
            }

            // Checking others
            DateTime CurrentDateTime = DateTime.Now;
            if (DavidBenoitEmailTimeHourList.Contains(CurrentDateTime.Hour) && CurrentDateTime.Minute == 0 && CurrentDateTime.Second == 0)
            {
                try
                {
                    string DavidBenoitErrorMessage = _TaskRunnerBaseService.RunTaskForDavidBenoit();
                    if (!string.IsNullOrWhiteSpace(DavidBenoitErrorMessage))
                    {
                        _RichTextBoxStatus.AppendText("David Benoit issue: " + DavidBenoitErrorMessage);
                    }
                }
                catch (Exception ex)
                {
                    _RichTextBoxStatus.AppendText("David Benoit issue: " + ex.Message + (ex.InnerException == null ? "" : ex.InnerException.Message));

                    MailMessage mail = new MailMessage();

                    mail.To.Add("David.Benoit@Canada.ca");
                    mail.Bcc.Add("Charles.LeBlanc2@Canada.ca");

                    mail.From = new MailAddress("ec.pccsm-cssp.ec@canada.ca");
                    mail.IsBodyHtml = true;

                    SmtpClient myClient = new System.Net.Mail.SmtpClient();

                    myClient.Host = "atlantic-exgate.Atlantic.int.ec.gc.ca";

                    string subject = "David Benoit Issue from CSSPWebToolsTaskRunner";

                    StringBuilder msg = new StringBuilder();

                    msg.AppendLine("<h2>David Benoit Issue Email</h2>");
                    msg.AppendLine("<h4>Date of issue: " + DateTime.Now + "</h4>");
                    msg.AppendLine("<h4>Exception Message: " + ex.Message + "</h4>");
                    msg.AppendLine("<h4>Exception Inner Message: " + (ex.InnerException != null ? ex.InnerException.Message : "empty") + "</h4>");

                    msg.AppendLine(@"<br>");
                    msg.AppendLine(@"<p>Auto email from CSSPWebTools.</p>");

                    mail.Subject = subject;
                    mail.Body = msg.ToString();
                    myClient.Send(mail);
                }
            }
        }
        public string UpdateOtherServerWithOtherServerLabSheetIDAndLabSheetStatus(int OtherServerLabSheetID, LabSheetStatusEnum LabSheetStatus)
        {
            string retStr = "";

            NameValueCollection paramList = new NameValueCollection();
            paramList.Add("OtherServerLabSheetID", OtherServerLabSheetID.ToString());
            paramList.Add("LabSheetStatus", ((int)LabSheetStatus).ToString());

            using (WebClient webClient = new WebClient())
            {
                WebProxy webProxy = new WebProxy();
                webClient.Proxy = webProxy;

                webClient.Headers.Add("Content-Type", "application/x-www-form-urlencoded");
                byte[] ret = webClient.UploadValues(new Uri("http://cssplabsheet.azurewebsites.net/UpdateLabSheetStatus.aspx"), "POST", paramList);

                retStr = System.Text.Encoding.Default.GetString(ret);
            }

            return retStr;
        }
        public string UploadLabSheetDetailInDB(LabSheetModelAndA1Sheet labSheetModelAndA1Sheet)
        {
            string retStr = "";

            // Filling LabSheetDetailModel
            LabSheetDetailModel labSheetDetailModelNew = new LabSheetDetailModel();

            labSheetDetailModelNew.LabSheetID = labSheetModelAndA1Sheet.LabSheetModel.LabSheetID;
            labSheetDetailModelNew.SamplingPlanID = labSheetModelAndA1Sheet.LabSheetModel.SamplingPlanID;
            labSheetDetailModelNew.SubsectorTVItemID = labSheetModelAndA1Sheet.LabSheetModel.SubsectorTVItemID;
            labSheetDetailModelNew.Version = labSheetModelAndA1Sheet.LabSheetA1Sheet.Version;
            labSheetDetailModelNew.IncludeLaboratoryQAQC = labSheetModelAndA1Sheet.LabSheetA1Sheet.IncludeLaboratoryQAQC;

            // RunDate
            DateTime RunDate = new DateTime();
            int RunYear = (string.IsNullOrWhiteSpace(labSheetModelAndA1Sheet.LabSheetA1Sheet.RunYear) ? 1900 : int.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.RunYear));
            int RunMonth = (string.IsNullOrWhiteSpace(labSheetModelAndA1Sheet.LabSheetA1Sheet.RunMonth) ? 1 : int.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.RunMonth));
            int RunDay = (string.IsNullOrWhiteSpace(labSheetModelAndA1Sheet.LabSheetA1Sheet.RunDay) ? 1 : int.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.RunDay));
            if (RunYear != 1900)
            {
                RunDate = new DateTime(RunYear, RunMonth, RunDay);
            }

            labSheetDetailModelNew.RunDate = RunDate;

            labSheetDetailModelNew.Tides = labSheetModelAndA1Sheet.LabSheetA1Sheet.Tides;
            labSheetDetailModelNew.SampleCrewInitials = labSheetModelAndA1Sheet.LabSheetA1Sheet.SampleCrewInitials;

            if (labSheetModelAndA1Sheet.LabSheetA1Sheet.IncludeLaboratoryQAQC)
            {
                // IncubationBath1StartDate
                DateTime IncubationBath1StartDate = new DateTime();
                int IncubationBath1StartYear = (string.IsNullOrWhiteSpace(labSheetModelAndA1Sheet.LabSheetA1Sheet.RunYear) ? 1900 : int.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.RunYear));
                int IncubationBath1StartMonth = (string.IsNullOrWhiteSpace(labSheetModelAndA1Sheet.LabSheetA1Sheet.RunMonth) ? 1 : int.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.RunMonth));
                int IncubationBath1StartDay = (string.IsNullOrWhiteSpace(labSheetModelAndA1Sheet.LabSheetA1Sheet.RunDay) ? 1 : int.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.RunDay));
                int IncubationBath1StartHour = 0;
                int IncubationBath1StartMinute = 0;
                if (labSheetModelAndA1Sheet.LabSheetA1Sheet.IncubationBath1StartTime.Length == 5)
                {
                    IncubationBath1StartHour = int.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.IncubationBath1StartTime.Substring(0, 2));
                    IncubationBath1StartMinute = int.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.IncubationBath1StartTime.Substring(3, 2));
                }
                if (IncubationBath1StartYear != 1900)
                {
                    IncubationBath1StartDate = new DateTime(IncubationBath1StartYear, IncubationBath1StartMonth, IncubationBath1StartDay, IncubationBath1StartHour, IncubationBath1StartMinute, 0);
                }

                labSheetDetailModelNew.IncubationBath1StartTime = IncubationBath1StartDate;

                // IncubationBath1EndDate
                DateTime IncubationBath1EndDate = new DateTime();
                int IncubationBath1EndYear = (string.IsNullOrWhiteSpace(labSheetModelAndA1Sheet.LabSheetA1Sheet.RunYear) ? 1900 : int.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.RunYear));
                int IncubationBath1EndMonth = (string.IsNullOrWhiteSpace(labSheetModelAndA1Sheet.LabSheetA1Sheet.RunMonth) ? 1 : int.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.RunMonth));
                int IncubationBath1EndDay = (string.IsNullOrWhiteSpace(labSheetModelAndA1Sheet.LabSheetA1Sheet.RunDay) ? 1 : int.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.RunDay));
                int IncubationBath1EndHour = 0;
                int IncubationBath1EndMinute = 0;
                if (labSheetModelAndA1Sheet.LabSheetA1Sheet.IncubationBath1EndTime.Length == 5)
                {
                    IncubationBath1EndHour = int.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.IncubationBath1EndTime.Substring(0, 2));
                    IncubationBath1EndMinute = int.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.IncubationBath1EndTime.Substring(3, 2));
                }
                if (IncubationBath1EndYear != 1900)
                {
                    IncubationBath1EndDate = new DateTime(IncubationBath1EndYear, IncubationBath1EndMonth, IncubationBath1EndDay, IncubationBath1EndHour, IncubationBath1EndMinute, 0).AddDays(1);
                }

                labSheetDetailModelNew.IncubationBath1EndTime = IncubationBath1EndDate;

                // IncubationBath1TimeCalculated_minutes 
                int IncubationBath1TimeCalculated_minutes = 0;
                if (labSheetModelAndA1Sheet.LabSheetA1Sheet.IncubationBath1TimeCalculated.Length == 5)
                {
                    IncubationBath1TimeCalculated_minutes = int.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.IncubationBath1TimeCalculated.Substring(0, 2)) * 60;
                    IncubationBath1TimeCalculated_minutes += int.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.IncubationBath1TimeCalculated.Substring(3, 2));
                }

                labSheetDetailModelNew.IncubationBath1TimeCalculated_minutes = IncubationBath1TimeCalculated_minutes;

                labSheetDetailModelNew.WaterBath1 = labSheetModelAndA1Sheet.LabSheetA1Sheet.WaterBath1;

                labSheetDetailModelNew.WaterBathCount = labSheetModelAndA1Sheet.LabSheetA1Sheet.WaterBathCount;

                if (labSheetDetailModelNew.WaterBathCount > 1)
                {
                    // IncubationBath2StartDate
                    DateTime IncubationBath2StartDate = new DateTime();
                    int IncubationBath2StartYear = (string.IsNullOrWhiteSpace(labSheetModelAndA1Sheet.LabSheetA1Sheet.RunYear) ? 1900 : int.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.RunYear));
                    int IncubationBath2StartMonth = (string.IsNullOrWhiteSpace(labSheetModelAndA1Sheet.LabSheetA1Sheet.RunMonth) ? 1 : int.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.RunMonth));
                    int IncubationBath2StartDay = (string.IsNullOrWhiteSpace(labSheetModelAndA1Sheet.LabSheetA1Sheet.RunDay) ? 1 : int.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.RunDay));
                    int IncubationBath2StartHour = 0;
                    int IncubationBath2StartMinute = 0;
                    if (labSheetModelAndA1Sheet.LabSheetA1Sheet.IncubationBath2StartTime.Length == 5)
                    {
                        IncubationBath2StartHour = int.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.IncubationBath2StartTime.Substring(0, 2));
                        IncubationBath2StartMinute = int.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.IncubationBath2StartTime.Substring(3, 2));
                    }
                    if (IncubationBath2StartYear != 1900)
                    {
                        IncubationBath2StartDate = new DateTime(IncubationBath2StartYear, IncubationBath2StartMonth, IncubationBath2StartDay, IncubationBath2StartHour, IncubationBath2StartMinute, 0);
                    }

                    labSheetDetailModelNew.IncubationBath2StartTime = IncubationBath2StartDate;

                    // IncubationBath2EndDate
                    DateTime IncubationBath2EndDate = new DateTime();
                    int IncubationBath2EndYear = (string.IsNullOrWhiteSpace(labSheetModelAndA1Sheet.LabSheetA1Sheet.RunYear) ? 1900 : int.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.RunYear));
                    int IncubationBath2EndMonth = (string.IsNullOrWhiteSpace(labSheetModelAndA1Sheet.LabSheetA1Sheet.RunMonth) ? 1 : int.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.RunMonth));
                    int IncubationBath2EndDay = (string.IsNullOrWhiteSpace(labSheetModelAndA1Sheet.LabSheetA1Sheet.RunDay) ? 1 : int.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.RunDay));
                    int IncubationBath2EndHour = 0;
                    int IncubationBath2EndMinute = 0;
                    if (labSheetModelAndA1Sheet.LabSheetA1Sheet.IncubationBath2EndTime.Length == 5)
                    {
                        IncubationBath2EndHour = int.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.IncubationBath2EndTime.Substring(0, 2));
                        IncubationBath2EndMinute = int.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.IncubationBath2EndTime.Substring(3, 2));
                    }
                    if (IncubationBath2EndYear != 1900)
                    {
                        IncubationBath2EndDate = new DateTime(IncubationBath2EndYear, IncubationBath2EndMonth, IncubationBath2EndDay, IncubationBath2EndHour, IncubationBath2EndMinute, 0).AddDays(1);
                    }

                    labSheetDetailModelNew.IncubationBath2EndTime = IncubationBath2EndDate;

                    // IncubationBath2TimeCalculated_minutes 
                    int IncubationBath2TimeCalculated_minutes = 0;
                    if (labSheetModelAndA1Sheet.LabSheetA1Sheet.IncubationBath2TimeCalculated.Length == 5)
                    {
                        IncubationBath2TimeCalculated_minutes = int.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.IncubationBath2TimeCalculated.Substring(0, 2)) * 60;
                        IncubationBath2TimeCalculated_minutes += int.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.IncubationBath2TimeCalculated.Substring(3, 2));
                    }

                    labSheetDetailModelNew.IncubationBath2TimeCalculated_minutes = IncubationBath2TimeCalculated_minutes;

                    labSheetDetailModelNew.WaterBath2 = labSheetModelAndA1Sheet.LabSheetA1Sheet.WaterBath2;
                    labSheetDetailModelNew.Bath2Positive44_5 = labSheetModelAndA1Sheet.LabSheetA1Sheet.Bath2Positive44_5;
                    labSheetDetailModelNew.Bath2NonTarget44_5 = labSheetModelAndA1Sheet.LabSheetA1Sheet.Bath2NonTarget44_5;
                    labSheetDetailModelNew.Bath2Negative44_5 = labSheetModelAndA1Sheet.LabSheetA1Sheet.Bath2Negative44_5;
                    labSheetDetailModelNew.Bath2Blank44_5 = labSheetModelAndA1Sheet.LabSheetA1Sheet.Bath2Blank44_5;

                }

                if (labSheetDetailModelNew.WaterBathCount > 2)
                {
                    // IncubationBath3StartDate
                    DateTime IncubationBath3StartDate = new DateTime();
                    int IncubationBath3StartYear = (string.IsNullOrWhiteSpace(labSheetModelAndA1Sheet.LabSheetA1Sheet.RunYear) ? 1900 : int.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.RunYear));
                    int IncubationBath3StartMonth = (string.IsNullOrWhiteSpace(labSheetModelAndA1Sheet.LabSheetA1Sheet.RunMonth) ? 1 : int.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.RunMonth));
                    int IncubationBath3StartDay = (string.IsNullOrWhiteSpace(labSheetModelAndA1Sheet.LabSheetA1Sheet.RunDay) ? 1 : int.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.RunDay));
                    int IncubationBath3StartHour = 0;
                    int IncubationBath3StartMinute = 0;
                    if (labSheetModelAndA1Sheet.LabSheetA1Sheet.IncubationBath3StartTime.Length == 5)
                    {
                        IncubationBath3StartHour = int.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.IncubationBath3StartTime.Substring(0, 2));
                        IncubationBath3StartMinute = int.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.IncubationBath3StartTime.Substring(3, 2));
                    }
                    if (IncubationBath3StartYear != 1900)
                    {
                        IncubationBath3StartDate = new DateTime(IncubationBath3StartYear, IncubationBath3StartMonth, IncubationBath3StartDay, IncubationBath3StartHour, IncubationBath3StartMinute, 0);
                    }

                    labSheetDetailModelNew.IncubationBath3StartTime = IncubationBath3StartDate;

                    // IncubationBath3EndDate
                    DateTime IncubationBath3EndDate = new DateTime();
                    int IncubationBath3EndYear = (string.IsNullOrWhiteSpace(labSheetModelAndA1Sheet.LabSheetA1Sheet.RunYear) ? 1900 : int.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.RunYear));
                    int IncubationBath3EndMonth = (string.IsNullOrWhiteSpace(labSheetModelAndA1Sheet.LabSheetA1Sheet.RunMonth) ? 1 : int.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.RunMonth));
                    int IncubationBath3EndDay = (string.IsNullOrWhiteSpace(labSheetModelAndA1Sheet.LabSheetA1Sheet.RunDay) ? 1 : int.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.RunDay));
                    int IncubationBath3EndHour = 0;
                    int IncubationBath3EndMinute = 0;
                    if (labSheetModelAndA1Sheet.LabSheetA1Sheet.IncubationBath3EndTime.Length == 5)
                    {
                        IncubationBath3EndHour = int.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.IncubationBath3EndTime.Substring(0, 2));
                        IncubationBath3EndMinute = int.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.IncubationBath3EndTime.Substring(3, 2));
                    }
                    if (IncubationBath3EndYear != 1900)
                    {
                        IncubationBath3EndDate = new DateTime(IncubationBath3EndYear, IncubationBath3EndMonth, IncubationBath3EndDay, IncubationBath3EndHour, IncubationBath3EndMinute, 0).AddDays(1);
                    }

                    labSheetDetailModelNew.IncubationBath3EndTime = IncubationBath3EndDate;

                    // IncubationBath3TimeCalculated_minutes 
                    int IncubationBath3TimeCalculated_minutes = 0;
                    if (labSheetModelAndA1Sheet.LabSheetA1Sheet.IncubationBath3TimeCalculated.Length == 5)
                    {
                        IncubationBath3TimeCalculated_minutes = int.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.IncubationBath3TimeCalculated.Substring(0, 2)) * 60;
                        IncubationBath3TimeCalculated_minutes += int.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.IncubationBath3TimeCalculated.Substring(3, 2));
                    }

                    labSheetDetailModelNew.IncubationBath3TimeCalculated_minutes = IncubationBath3TimeCalculated_minutes;

                    labSheetDetailModelNew.WaterBath3 = labSheetModelAndA1Sheet.LabSheetA1Sheet.WaterBath3;
                    labSheetDetailModelNew.Bath3Positive44_5 = labSheetModelAndA1Sheet.LabSheetA1Sheet.Bath3Positive44_5;
                    labSheetDetailModelNew.Bath3NonTarget44_5 = labSheetModelAndA1Sheet.LabSheetA1Sheet.Bath3NonTarget44_5;
                    labSheetDetailModelNew.Bath3Negative44_5 = labSheetModelAndA1Sheet.LabSheetA1Sheet.Bath3Negative44_5;
                    labSheetDetailModelNew.Bath3Blank44_5 = labSheetModelAndA1Sheet.LabSheetA1Sheet.Bath3Blank44_5;

                }

                labSheetDetailModelNew.TCField1 = null;
                if (!string.IsNullOrWhiteSpace(labSheetModelAndA1Sheet.LabSheetA1Sheet.TCField1))
                {
                    float temp = 0.0f;
                    if (float.TryParse(labSheetModelAndA1Sheet.LabSheetA1Sheet.TCField1, out temp))
                    {
                        labSheetDetailModelNew.TCField1 = temp;
                    }
                }
                labSheetDetailModelNew.TCLab1 = null;
                if (!string.IsNullOrWhiteSpace(labSheetModelAndA1Sheet.LabSheetA1Sheet.TCLab1))
                {
                    float temp = 0.0f;
                    if (float.TryParse(labSheetModelAndA1Sheet.LabSheetA1Sheet.TCLab1, out temp))
                    {
                        labSheetDetailModelNew.TCLab1 = temp;
                    }
                }
                labSheetDetailModelNew.TCField2 = null;
                if (!string.IsNullOrWhiteSpace(labSheetModelAndA1Sheet.LabSheetA1Sheet.TCField2))
                {
                    float temp = 0.0f;
                    if (float.TryParse(labSheetModelAndA1Sheet.LabSheetA1Sheet.TCField2, out temp))
                    {
                        labSheetDetailModelNew.TCField2 = temp;
                    }
                }
                labSheetDetailModelNew.TCLab2 = null;
                if (!string.IsNullOrWhiteSpace(labSheetModelAndA1Sheet.LabSheetA1Sheet.TCLab2))
                {
                    float temp = 0.0f;
                    if (float.TryParse(labSheetModelAndA1Sheet.LabSheetA1Sheet.TCLab2, out temp))
                    {
                        labSheetDetailModelNew.TCLab2 = temp;
                    }
                }
                labSheetDetailModelNew.TCFirst = null;
                if (!string.IsNullOrWhiteSpace(labSheetModelAndA1Sheet.LabSheetA1Sheet.TCFirst) && !labSheetModelAndA1Sheet.LabSheetA1Sheet.TCFirst.Contains("-"))
                {
                    float temp = 0.0f;
                    if (float.TryParse(labSheetModelAndA1Sheet.LabSheetA1Sheet.TCFirst, out temp))
                    {
                        labSheetDetailModelNew.TCFirst = temp;
                    }
                }
                labSheetDetailModelNew.TCAverage = null;
                if (!string.IsNullOrWhiteSpace(labSheetModelAndA1Sheet.LabSheetA1Sheet.TCAverage) && !labSheetModelAndA1Sheet.LabSheetA1Sheet.TCAverage.Contains("-"))
                {
                    float temp = 0.0f;
                    if (float.TryParse(labSheetModelAndA1Sheet.LabSheetA1Sheet.TCAverage, out temp))
                    {
                        labSheetDetailModelNew.TCAverage = temp;
                    }
                }
                labSheetDetailModelNew.ControlLot = labSheetModelAndA1Sheet.LabSheetA1Sheet.ControlLot;
                labSheetDetailModelNew.Positive35 = labSheetModelAndA1Sheet.LabSheetA1Sheet.Positive35;
                labSheetDetailModelNew.NonTarget35 = labSheetModelAndA1Sheet.LabSheetA1Sheet.NonTarget35;
                labSheetDetailModelNew.Negative35 = labSheetModelAndA1Sheet.LabSheetA1Sheet.Negative35;
                labSheetDetailModelNew.Bath1Positive44_5 = labSheetModelAndA1Sheet.LabSheetA1Sheet.Bath1Positive44_5;
                labSheetDetailModelNew.Bath1NonTarget44_5 = labSheetModelAndA1Sheet.LabSheetA1Sheet.Bath1NonTarget44_5;
                labSheetDetailModelNew.Bath1Negative44_5 = labSheetModelAndA1Sheet.LabSheetA1Sheet.Bath1Negative44_5;
                labSheetDetailModelNew.Blank35 = labSheetModelAndA1Sheet.LabSheetA1Sheet.Blank35;
                labSheetDetailModelNew.Bath1Blank44_5 = labSheetModelAndA1Sheet.LabSheetA1Sheet.Bath1Blank44_5;
                labSheetDetailModelNew.Lot35 = labSheetModelAndA1Sheet.LabSheetA1Sheet.Lot35;
                labSheetDetailModelNew.Lot44_5 = labSheetModelAndA1Sheet.LabSheetA1Sheet.Lot44_5;
                labSheetDetailModelNew.RunComment = labSheetModelAndA1Sheet.LabSheetA1Sheet.RunComment;
                labSheetDetailModelNew.RunWeatherComment = labSheetModelAndA1Sheet.LabSheetA1Sheet.RunWeatherComment;
                labSheetDetailModelNew.SampleBottleLotNumber = labSheetModelAndA1Sheet.LabSheetA1Sheet.SampleBottleLotNumber;
                labSheetDetailModelNew.SalinitiesReadBy = labSheetModelAndA1Sheet.LabSheetA1Sheet.SalinitiesReadBy;

                // SalinitiesReadDate
                DateTime SalinitiesReadDate = new DateTime();
                int SalinitiesReadYear = (string.IsNullOrWhiteSpace(labSheetModelAndA1Sheet.LabSheetA1Sheet.SalinitiesReadYear) ? 1900 : int.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.SalinitiesReadYear));
                int SalinitiesReadMonth = (string.IsNullOrWhiteSpace(labSheetModelAndA1Sheet.LabSheetA1Sheet.SalinitiesReadMonth) ? 1 : int.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.SalinitiesReadMonth));
                int SalinitiesReadDay = (string.IsNullOrWhiteSpace(labSheetModelAndA1Sheet.LabSheetA1Sheet.SalinitiesReadDay) ? 1 : int.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.SalinitiesReadDay));
                if (SalinitiesReadYear != 1900)
                {
                    SalinitiesReadDate = new DateTime(SalinitiesReadYear, SalinitiesReadMonth, SalinitiesReadDay);
                }

                labSheetDetailModelNew.SalinitiesReadDate = SalinitiesReadDate;
                labSheetDetailModelNew.ResultsReadBy = labSheetModelAndA1Sheet.LabSheetA1Sheet.ResultsReadBy;

                // ResultsReadDate
                DateTime ResultsReadDate = new DateTime();
                int ResultsReadYear = (string.IsNullOrWhiteSpace(labSheetModelAndA1Sheet.LabSheetA1Sheet.ResultsReadYear) ? 1900 : int.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.ResultsReadYear));
                int ResultsReadMonth = (string.IsNullOrWhiteSpace(labSheetModelAndA1Sheet.LabSheetA1Sheet.ResultsReadMonth) ? 1 : int.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.ResultsReadMonth));
                int ResultsReadDay = (string.IsNullOrWhiteSpace(labSheetModelAndA1Sheet.LabSheetA1Sheet.ResultsReadDay) ? 1 : int.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.ResultsReadDay));
                if (ResultsReadYear != 1900)
                {
                    ResultsReadDate = new DateTime(ResultsReadYear, ResultsReadMonth, ResultsReadDay);
                }

                labSheetDetailModelNew.ResultsReadDate = ResultsReadDate;
                labSheetDetailModelNew.ResultsRecordedBy = labSheetModelAndA1Sheet.LabSheetA1Sheet.ResultsRecordedBy;

                // ResultsRecordedDate
                DateTime ResultsRecordedDate = new DateTime();
                int ResultsRecordedYear = (string.IsNullOrWhiteSpace(labSheetModelAndA1Sheet.LabSheetA1Sheet.ResultsRecordedYear) ? 1900 : int.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.ResultsRecordedYear));
                int ResultsRecordedMonth = (string.IsNullOrWhiteSpace(labSheetModelAndA1Sheet.LabSheetA1Sheet.ResultsRecordedMonth) ? 1 : int.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.ResultsRecordedMonth));
                int ResultsRecordedDay = (string.IsNullOrWhiteSpace(labSheetModelAndA1Sheet.LabSheetA1Sheet.ResultsRecordedDay) ? 1 : int.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.ResultsRecordedDay));
                if (ResultsRecordedYear != 1900)
                {
                    ResultsRecordedDate = new DateTime(ResultsRecordedYear, ResultsRecordedMonth, ResultsRecordedDay);
                }

                labSheetDetailModelNew.ResultsRecordedDate = ResultsRecordedDate;

                labSheetDetailModelNew.DailyDuplicateRLog = 0.0f;
                if (!string.IsNullOrWhiteSpace(labSheetModelAndA1Sheet.LabSheetA1Sheet.DailyDuplicateRLog) && !labSheetModelAndA1Sheet.LabSheetA1Sheet.DailyDuplicateRLog.StartsWith("N"))
                {
                    try
                    {
                        labSheetDetailModelNew.DailyDuplicateRLog = float.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.DailyDuplicateRLog);
                    }
                    catch (Exception)
                    {
                        // nothing
                    }
                }

                labSheetDetailModelNew.DailyDuplicatePrecisionCriteria = null;
                labSheetDetailModelNew.DailyDuplicateRLog = null;
                labSheetDetailModelNew.DailyDuplicateAcceptable = null;
                labSheetDetailModelNew.IntertechDuplicatePrecisionCriteria = null;
                labSheetDetailModelNew.IntertechDuplicateRLog = null;
                labSheetDetailModelNew.IntertechDuplicateAcceptable = null;
                labSheetDetailModelNew.IntertechReadAcceptable = null;

                if (!string.IsNullOrWhiteSpace(labSheetModelAndA1Sheet.LabSheetA1Sheet.DailyDuplicatePrecisionCriteria))
                {
                    labSheetDetailModelNew.DailyDuplicatePrecisionCriteria = float.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.DailyDuplicatePrecisionCriteria);
                }

                if (!string.IsNullOrWhiteSpace(labSheetModelAndA1Sheet.LabSheetA1Sheet.DailyDuplicateRLog))
                {
                    if (!labSheetModelAndA1Sheet.LabSheetA1Sheet.DailyDuplicateRLog.StartsWith("N"))
                    {
                        labSheetDetailModelNew.DailyDuplicateRLog = float.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.DailyDuplicateRLog);
                    }
                }

                if (!string.IsNullOrWhiteSpace(labSheetModelAndA1Sheet.LabSheetA1Sheet.DailyDuplicateAcceptableOrUnacceptable))
                {
                    labSheetDetailModelNew.DailyDuplicateAcceptable = (labSheetModelAndA1Sheet.LabSheetA1Sheet.DailyDuplicateAcceptableOrUnacceptable != "Acceptable" ? false : true);
                }

                if (!string.IsNullOrWhiteSpace(labSheetModelAndA1Sheet.LabSheetA1Sheet.IntertechDuplicatePrecisionCriteria))
                {
                    labSheetDetailModelNew.IntertechDuplicatePrecisionCriteria = float.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.IntertechDuplicatePrecisionCriteria);
                }

                if (!string.IsNullOrWhiteSpace(labSheetModelAndA1Sheet.LabSheetA1Sheet.IntertechDuplicateRLog))
                {
                    if (!labSheetModelAndA1Sheet.LabSheetA1Sheet.IntertechDuplicateRLog.StartsWith("N"))
                    {
                        labSheetDetailModelNew.IntertechDuplicateRLog = float.Parse(labSheetModelAndA1Sheet.LabSheetA1Sheet.IntertechDuplicateRLog);
                    }
                }

                if (!string.IsNullOrWhiteSpace(labSheetModelAndA1Sheet.LabSheetA1Sheet.IntertechDuplicateAcceptableOrUnacceptable))
                {
                    labSheetDetailModelNew.IntertechDuplicateAcceptable = (labSheetModelAndA1Sheet.LabSheetA1Sheet.IntertechDuplicateAcceptableOrUnacceptable != "Acceptable" ? false : true);
                }

                if (!string.IsNullOrWhiteSpace(labSheetModelAndA1Sheet.LabSheetA1Sheet.IntertechReadAcceptableOrUnacceptable))
                {
                    labSheetDetailModelNew.IntertechReadAcceptable = (labSheetModelAndA1Sheet.LabSheetA1Sheet.IntertechReadAcceptableOrUnacceptable != "Acceptable" ? false : true);
                }
            }

            LabSheetDetailService labSheetDetailService = new LabSheetDetailService(LanguageEnum.en, _User);

            LabSheetDetailModel labSheetDetailModelExist = labSheetDetailService.GetLabSheetDetailModelExistDB(labSheetDetailModelNew);
            if (!string.IsNullOrWhiteSpace(labSheetDetailModelExist.Error))
            {
                labSheetDetailModelExist = labSheetDetailService.PostAddLabSheetDetailDB(labSheetDetailModelNew);
                if (!string.IsNullOrWhiteSpace(labSheetDetailModelExist.Error))
                {
                    richTextBoxStatus.AppendText("Lab sheet detail could not be loaded to the local DB. Error [" + labSheetDetailModelExist.Error + "]\r\n");
                    return labSheetDetailModelExist.Error;
                }
            }
            else
            {
                labSheetDetailModelNew.LabSheetDetailID = labSheetDetailModelExist.LabSheetDetailID;
                labSheetDetailModelExist = labSheetDetailService.PostUpdateLabSheetDetailDB(labSheetDetailModelNew);
                if (!string.IsNullOrWhiteSpace(labSheetDetailModelExist.Error))
                {
                    richTextBoxStatus.AppendText("Lab sheet detail could not be loaded to the local DB. Error [" + labSheetDetailModelExist.Error + "]\r\n");
                    return labSheetDetailModelExist.Error;
                }
            }

            retStr = UploadLabSheetTubeMPNDetailInDB(labSheetDetailModelExist.LabSheetDetailID, labSheetModelAndA1Sheet.LabSheetA1Sheet.LabSheetA1MeasurementList);
            if (!string.IsNullOrWhiteSpace(retStr))
            {
                richTextBoxStatus.AppendText("Lab sheet tube and MPN detail could not be loaded to the local DB. Error [" + retStr + "]\r\n");
                return retStr;
            }

            return retStr;
        }
        public string UploadLabSheetTubeMPNDetailInDB(int LabSheetDetailID, List<LabSheetA1Measurement> labSheetA1MeasurementList)
        {
            string retStr = "";

            LabSheetTubeMPNDetailService labSheetTubeMPNDetailService = new LabSheetTubeMPNDetailService(LanguageEnum.en, _User);

            int Ordinal = 0;
            foreach (LabSheetA1Measurement labSheetA1Measurement in labSheetA1MeasurementList)
            {
                LabSheetTubeMPNDetailModel labSheetTubeMPNDetailModelNew = new LabSheetTubeMPNDetailModel();
                labSheetTubeMPNDetailModelNew.LabSheetDetailID = LabSheetDetailID;
                labSheetTubeMPNDetailModelNew.Ordinal = Ordinal;
                labSheetTubeMPNDetailModelNew.MWQMSiteTVItemID = labSheetA1Measurement.TVItemID;
                labSheetTubeMPNDetailModelNew.SampleDateTime = labSheetA1Measurement.Time;
                labSheetTubeMPNDetailModelNew.MPN = labSheetA1Measurement.MPN;
                labSheetTubeMPNDetailModelNew.Tube10 = labSheetA1Measurement.Tube10;
                labSheetTubeMPNDetailModelNew.Tube1_0 = labSheetA1Measurement.Tube1_0;
                labSheetTubeMPNDetailModelNew.Tube0_1 = labSheetA1Measurement.Tube0_1;
                labSheetTubeMPNDetailModelNew.Salinity = labSheetA1Measurement.Salinity;
                labSheetTubeMPNDetailModelNew.Temperature = labSheetA1Measurement.Temperature;
                labSheetTubeMPNDetailModelNew.ProcessedBy = labSheetA1Measurement.ProcessedBy;
                labSheetTubeMPNDetailModelNew.SampleType = (SampleTypeEnum)labSheetA1Measurement.SampleType;
                labSheetTubeMPNDetailModelNew.SiteComment = labSheetA1Measurement.SiteComment;

                LabSheetTubeMPNDetailModel labSheetTubeMPNDetailModelExist = labSheetTubeMPNDetailService.GetLabSheetTubeMPNDetailModelExistDB(labSheetTubeMPNDetailModelNew);
                if (!string.IsNullOrWhiteSpace(labSheetTubeMPNDetailModelExist.Error))
                {
                    labSheetTubeMPNDetailModelExist = labSheetTubeMPNDetailService.PostAddLabSheetTubeMPNDetailDB(labSheetTubeMPNDetailModelNew);
                    if (!string.IsNullOrWhiteSpace(labSheetTubeMPNDetailModelExist.Error))
                        return labSheetTubeMPNDetailModelExist.Error;
                }
                else
                {
                    labSheetTubeMPNDetailModelNew.LabSheetTubeMPNDetailID = labSheetTubeMPNDetailModelExist.LabSheetTubeMPNDetailID;
                    labSheetTubeMPNDetailModelExist = labSheetTubeMPNDetailService.PostUpdateLabSheetTubeMPNDetailDB(labSheetTubeMPNDetailModelNew);
                    if (!string.IsNullOrWhiteSpace(labSheetTubeMPNDetailModelExist.Error))
                        return labSheetTubeMPNDetailModelExist.Error;
                }

                Ordinal += 1;
            }

            return retStr;
        }
        public void GetNextLabSheet()
        {
            string retStr = "";
            try
            {
                using (WebClient webClient = new WebClient())
                {
                    WebProxy webProxy = new WebProxy();
                    webClient.Proxy = webProxy;

                    string FullLabSheetText = webClient.DownloadString(new Uri("http://cssplabsheet.azurewebsites.net/GetNextLabSheet.aspx"));

                    if (FullLabSheetText.Length > 0)
                    {
                        int posStart = FullLabSheetText.IndexOf("OtherServerLabSheetID|||||[") + 27;
                        int posEnd = FullLabSheetText.IndexOf("]", posStart);
                        string OtherServerLabSheetIDTxt = FullLabSheetText.Substring(posStart, posEnd - posStart);
                        int OtherServerLabSheetID = int.Parse(OtherServerLabSheetIDTxt);

                        posStart = FullLabSheetText.IndexOf("SubsectorTVItemID|||||[") + 23;
                        posEnd = FullLabSheetText.IndexOf("]", posStart);
                        string SubsectorTVItemIDTxt = FullLabSheetText.Substring(posStart, posEnd - posStart);
                        int SubsectorTVItemID = int.Parse(SubsectorTVItemIDTxt);

                        TVItemModel tvItemModelSubsector = new TVItemModel();
                        TVItemModel tvItemModelCountry = new TVItemModel();
                        TVItemModel tvItemModelProvince = new TVItemModel();
                        LabSheetModelAndA1Sheet labSheetModelAndA1Sheet = new LabSheetModelAndA1Sheet();
                        using (TransactionScope ts = new TransactionScope())
                        {
                            LabSheetService labSheetService = new LabSheetService(LanguageEnum.en, _User);

                            LabSheetModel labSheetModelRet = labSheetService.AddOrUpdateLabSheetDB(FullLabSheetText);
                            if (!string.IsNullOrWhiteSpace(labSheetModelRet.Error))
                            {
                                richTextBoxStatus.AppendText("Lab sheet adding error OtherServerLabSheetID [" + OtherServerLabSheetID.ToString() + "]" + labSheetModelRet.Error + "]\r\n");
                                return;
                            }

                            TVItemService tvItemService = new TVItemService(LanguageEnum.en, _User);

                            tvItemModelSubsector = tvItemService.GetTVItemModelWithTVItemIDDB(SubsectorTVItemID);
                            if (!string.IsNullOrWhiteSpace(tvItemModelSubsector.Error))
                            {
                                richTextBoxStatus.AppendText("Lab sheet parsing error OtherServerLabSheetID [" + OtherServerLabSheetID.ToString() + "]" + tvItemModelSubsector.Error + "]\r\n");
                                return;
                            }

                            List<TVItemModel> tvItemModelList = tvItemService.GetParentsTVItemModelList(tvItemModelSubsector.TVPath);
                            foreach (TVItemModel tvItemModel in tvItemModelList)
                            {
                                if (tvItemModel.TVType == TVTypeEnum.Province)
                                {
                                    tvItemModelProvince = tvItemModel;
                                }
                                if (tvItemModel.TVType == TVTypeEnum.Country)
                                {
                                    tvItemModelCountry = tvItemModel;
                                }
                            }

                            labSheetModelAndA1Sheet.LabSheetModel = labSheetModelRet;
                            labSheetModelAndA1Sheet.LabSheetA1Sheet = labSheetService.ParseLabSheetA1WithLabSheetID(labSheetModelRet.LabSheetID);
                            if (!string.IsNullOrWhiteSpace(labSheetModelAndA1Sheet.LabSheetA1Sheet.Error))
                            {
                                richTextBoxStatus.AppendText("Lab sheet parsing error OtherServerLabSheetID [" + OtherServerLabSheetID.ToString() + "]" + labSheetModelAndA1Sheet.LabSheetA1Sheet.Error + "]\r\n");
                                richTextBoxStatus.AppendText("Full Lab Sheet Text below\r\n");
                                richTextBoxStatus.AppendText("---------------- Start of full lab sheet text -----------\r\n");
                                richTextBoxStatus.AppendText(FullLabSheetText);
                                richTextBoxStatus.AppendText("---------------- End of full lab sheet text -----------\r\n");
                                retStr = UpdateOtherServerWithOtherServerLabSheetIDAndLabSheetStatus(OtherServerLabSheetID, LabSheetStatusEnum.Error);
                                if (!string.IsNullOrWhiteSpace(retStr))
                                {
                                    richTextBoxStatus.AppendText("Error updating other server lab sheet [" + retStr + "]");
                                }
                                return;
                            }

                            string retStr2 = UploadLabSheetDetailInDB(labSheetModelAndA1Sheet);
                            if (!string.IsNullOrWhiteSpace(retStr2))
                            {
                                // Error message already sent to richTextboxStatus
                                retStr = UpdateOtherServerWithOtherServerLabSheetIDAndLabSheetStatus(OtherServerLabSheetID, LabSheetStatusEnum.Error);
                                if (!string.IsNullOrWhiteSpace(retStr))
                                {
                                    richTextBoxStatus.AppendText("Error updating other server lab sheet [" + retStr + "]");
                                }
                                return;
                            }

                            ts.Complete();
                        }

                        string href = "http://wmon01dtchlebl2/csspwebtools/en-CA/#!View/" + (tvItemModelCountry.TVText + "-" + tvItemModelProvince.TVText).Replace(" ", "-") + "|||" + tvItemModelProvince.TVItemID.ToString() + "|||010003030200000000000000000000";

                        if (labSheetModelAndA1Sheet.LabSheetA1Sheet.LabSheetA1MeasurementList.Where(c => c.MPN != null && c.MPN >= MPNLimitForEmail).Any())
                        {
                            SendNewLabSheetEmailBigMPN(href, tvItemModelProvince, tvItemModelSubsector, labSheetModelAndA1Sheet);
                        }
                        else
                        {
                            SendNewLabSheetEmail(href, tvItemModelProvince, tvItemModelSubsector, labSheetModelAndA1Sheet);
                        }

                        retStr = UpdateOtherServerWithOtherServerLabSheetIDAndLabSheetStatus(OtherServerLabSheetID, LabSheetStatusEnum.Transferred);
                        if (!string.IsNullOrWhiteSpace(retStr))
                        {
                            richTextBoxStatus.AppendText("Error updating other server lab sheet [" + retStr + "]");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                string errrrrr = ex.Message;
            }
        }
        private void SendNewLabSheetEmailBigMPN(string href, TVItemModel tvItemModelProvince, TVItemModel tvItemModelSubsector, LabSheetModelAndA1Sheet labSheetModelAndA1Sheet)
        {
            ContactService contactService = new ContactService(LanguageEnum.en, _User);

            List<ContactModel> contactModelList = contactService.GetContactModelWithSamplingPlanner_ProvincesTVItemIDDB(tvItemModelProvince.TVItemID);

            MailMessage mail = new MailMessage();

            if (!testing)
            {
                foreach (ContactModel contactModel in contactModelList)
                {
                    mail.To.Add(contactModel.LoginEmail.ToLower());
                }
            }
            else
            {
                mail.To.Add("charles.leblanc2@Canada.ca");
            }

            mail.From = new MailAddress("ec.pccsm-cssp.ec@canada.ca");
            mail.IsBodyHtml = true;

            SmtpClient myClient = new System.Net.Mail.SmtpClient();

            myClient.Host = "atlantic-exgate.Atlantic.int.ec.gc.ca";
            mail.Priority = MailPriority.High;

            string subject = "High MPN - Lab Sheets Sent --- " + tvItemModelProvince.TVText + " --- " + tvItemModelSubsector.TVText;

            StringBuilder msg = new StringBuilder();
            msg.AppendLine(@"<h2>New Lab Sheet Sent for " + tvItemModelProvince.TVText + "</h2>");
            msg.AppendLine(@"<h4>Subsector: " + tvItemModelSubsector.TVText + "</h4>");
            msg.AppendLine(@"<h4>Run Date: " + labSheetModelAndA1Sheet.LabSheetA1Sheet.RunYear + "-" +
                (labSheetModelAndA1Sheet.LabSheetA1Sheet.RunMonth.Length == 1 ? "0" + labSheetModelAndA1Sheet.LabSheetA1Sheet.RunMonth : labSheetModelAndA1Sheet.LabSheetA1Sheet.RunMonth) + "-" +
                (labSheetModelAndA1Sheet.LabSheetA1Sheet.RunDay.Length == 1 ? "0" + labSheetModelAndA1Sheet.LabSheetA1Sheet.RunDay : labSheetModelAndA1Sheet.LabSheetA1Sheet.RunDay) + "</h4>");
            msg.AppendLine(@"<a href=""" + href + @""">Open CSSPWebTools</a>");
            msg.AppendLine(@"<br /><br />");
            msg.AppendLine(@"<b>Note: </b> Sites below are over the MPN threshold of " + MPNLimitForEmail.ToString() + "<br />");

            List<LabSheetA1Measurement> labSheetA1MeasurementList = (from c in labSheetModelAndA1Sheet.LabSheetA1Sheet.LabSheetA1MeasurementList
                                                                     where c.MPN != null
                                                                     && c.MPN >= MPNLimitForEmail
                                                                     select c).ToList();

            msg.AppendLine("<ol>");
            foreach (LabSheetA1Measurement labSheetA1Measurment in labSheetA1MeasurementList)
            {
                msg.AppendLine("<li>");
                msg.AppendLine("<b>Site: </b>" + labSheetA1Measurment.Site + (labSheetA1Measurment.SampleType == SampleTypeEnum.DailyDuplicate ? " Daily duplicate" : "") + " <b>MPN: </b>" + labSheetA1Measurment.MPN);
                msg.AppendLine("</li>");

            }
            msg.AppendLine("</ol>");

            msg.AppendLine(@"<br>");
            msg.AppendLine(@"<p>Auto email from CSSPWebTools.</p>");

            mail.Subject = subject;
            mail.Body = msg.ToString();
            myClient.Send(mail);
        }
        public void SendNewLabSheetEmail(string href, TVItemModel tvItemModelProvince, TVItemModel tvItemModelSubsector, LabSheetModelAndA1Sheet labSheetModelAndA1Sheet)
        {
            ContactService contactService = new ContactService(LanguageEnum.en, _User);

            List<ContactModel> contactModelList = contactService.GetContactModelWithSamplingPlanner_ProvincesTVItemIDDB(tvItemModelProvince.TVItemID);

            MailMessage mail = new MailMessage();

            if (!testing)
            {
                foreach (ContactModel contactModel in contactModelList)
                {
                    mail.To.Add(contactModel.LoginEmail.ToLower());
                }
            }
            else
            {
                mail.To.Add("charles.leblanc2@Canada.ca");
            }

            mail.From = new MailAddress("ec.pccsm-cssp.ec@canada.ca");
            mail.IsBodyHtml = true;

            SmtpClient myClient = new System.Net.Mail.SmtpClient();

            myClient.Host = "atlantic-exgate.Atlantic.int.ec.gc.ca";

            string subject = "Lab Sheets Sent --- " + tvItemModelProvince.TVText + " --- " + tvItemModelSubsector.TVText;

            StringBuilder msg = new StringBuilder();

            msg.AppendLine("<h2>New Lab Sheet Sent for " + tvItemModelProvince.TVText + "</h2>");
            msg.AppendLine("<h4>Subsector: " + tvItemModelSubsector.TVText + "</h4>");
            msg.AppendLine(@"<h4>Run Date: " + labSheetModelAndA1Sheet.LabSheetA1Sheet.RunYear + "-" +
                (labSheetModelAndA1Sheet.LabSheetA1Sheet.RunMonth.Length == 1 ? "0" + labSheetModelAndA1Sheet.LabSheetA1Sheet.RunMonth : labSheetModelAndA1Sheet.LabSheetA1Sheet.RunMonth) + "-" +
                (labSheetModelAndA1Sheet.LabSheetA1Sheet.RunDay.Length == 1 ? "0" + labSheetModelAndA1Sheet.LabSheetA1Sheet.RunDay : labSheetModelAndA1Sheet.LabSheetA1Sheet.RunDay) + "</h4>");
            msg.AppendLine("<a href=\"" + href + "\">Open CSSPWebTools</a>");

            msg.AppendLine(@"<br>");
            msg.AppendLine(@"<p>Auto email from CSSPWebTools.</p>");

            mail.Subject = subject;
            mail.Body = msg.ToString();
            myClient.Send(mail);
        }
        public void StartTimer()
        {
            timerCheckTask.Enabled = true;
        }
        public void StopTimer()
        {
            timerCheckTask.Enabled = false;
        }
        public void WriteToRichTextBoxStatus(string str)
        {
            try
            {
                richTextBoxStatus.AppendText(str);
            }
            catch (Exception)
            {
                // nothing
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            FileInfo fiTemplate = new FileInfo(@"\\atlantic.int.ec.gc.ca\shares\Branches\EPB\ShellFish\All Contacts - MWQM Emergency, Spills, etc\EmailListTemplate.xlsx");
            if (!fiTemplate.Exists)
            {
                return;
            }

            IPrincipal UserTemp = new GenericPrincipal(new GenericIdentity("Charles.LeBlanc2@canada.ca", "Forms"), null);
            EmailDistributionListService emailDistributionListService = new EmailDistributionListService(LanguageEnum.en, UserTemp);
            EmailDistributionListContactService emailDistributionListContactService = new EmailDistributionListContactService(LanguageEnum.en, UserTemp);

            Microsoft.Office.Interop.Excel.Application appExcel = new Microsoft.Office.Interop.Excel.Application();
            appExcel.Visible = false;
            Microsoft.Office.Interop.Excel.Workbook xlWorkbook = appExcel.Workbooks.Open(fiTemplate.FullName);
            Microsoft.Office.Interop.Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];

            int StartRow = 143;
            int EndRow = 155;
            int Ordinal = 7;
            EmailDistributionListModel emailDistributionListModelRet = new EmailDistributionListModel();
            for (int i = StartRow; i <= EndRow; i++)
            {
                if (i == StartRow)
                {
                    EmailDistributionListModel emailDistributionListModelNew = new EmailDistributionListModel();

                    emailDistributionListModelNew.RegionName = xlWorksheet.Cells[i, 1].Value;
                    emailDistributionListModelNew.Ordinal = Ordinal;
                    emailDistributionListModelNew.CountryTVItemID = 5;

                    emailDistributionListModelRet = emailDistributionListService.PostAddEmailDistributionListDB(emailDistributionListModelNew);
                    if (!string.IsNullOrWhiteSpace(emailDistributionListModelRet.Error))
                    {
                        return;
                    }
                }

                EmailDistributionListContactModel emailDistributionListModelContactNew = new EmailDistributionListContactModel();

                emailDistributionListModelContactNew.EmailDistributionListID = emailDistributionListModelRet.EmailDistributionListID;
                emailDistributionListModelContactNew.IsCC = (xlWorksheet.Cells[i, 2].Value == "CC" ? true : false);
                emailDistributionListModelContactNew.Agency = xlWorksheet.Cells[i, 3].Value;
                emailDistributionListModelContactNew.Name = xlWorksheet.Cells[i, 4].Value;
                emailDistributionListModelContactNew.Email = xlWorksheet.Cells[i, 5].Value;
                emailDistributionListModelContactNew.CMPRainfallSeasonal = (xlWorksheet.Cells[i, 6].Value == "Y" ? true : false);
                emailDistributionListModelContactNew.CMPWastewater = (xlWorksheet.Cells[i, 7].Value == "Y" ? true : false);
                emailDistributionListModelContactNew.EmergencyWeather = (xlWorksheet.Cells[i, 8].Value == "Y" ? true : false);
                emailDistributionListModelContactNew.EmergencyWastewater = (xlWorksheet.Cells[i, 9].Value == "Y" ? true : false);
                emailDistributionListModelContactNew.ReopeningAllTypes = (xlWorksheet.Cells[i, 10].Value == "Y" ? true : false);

                EmailDistributionListContactModel emailDistributionListModelContactRet = emailDistributionListContactService.PostAddEmailDistributionListContactDB(emailDistributionListModelContactNew);
                if (!string.IsNullOrWhiteSpace(emailDistributionListModelContactRet.Error))
                {
                    return;
                }
            }

            xlWorkbook.Close();
            appExcel.Quit();

        }
        #endregion Functions
        //private string GetFileName(PFSFile pfsFile, string Path, string Keyword)
        //{
        //    string FileName = "";

        //    PFSSection pfsSectionFileName = pfsFile.GetSectionFromHandle(Path);

        //    if (pfsSectionFileName != null)
        //    {
        //        PFSKeyword keyword = null;
        //        try
        //        {
        //            keyword = pfsSectionFileName.GetKeyword(Keyword);
        //        }
        //        catch (Exception ex)
        //        {
        //            richTextBoxStatus.AppendText(ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
        //            return FileName;
        //        }

        //        if (keyword != null)
        //        {
        //            try
        //            {
        //                FileName = keyword.GetParameter(1).ToFileNamePath();
        //            }
        //            catch (Exception ex)
        //            {
        //                richTextBoxStatus.AppendText(ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
        //                return FileName;
        //            }
        //        }
        //    }

        //    return FileName;
        //}

        //private string SetFileName(PFSFile pfsFile, string Path, string Keyword, string NewFileName)
        //{
        //    string FileName = "";

        //    PFSSection pfsSectionFileName = pfsFile.GetSectionFromHandle(Path);

        //    if (pfsSectionFileName != null)
        //    {
        //        PFSKeyword keyword = null;
        //        try
        //        {
        //            keyword = pfsSectionFileName.GetKeyword(Keyword);
        //        }
        //        catch (Exception ex)
        //        {
        //            richTextBoxStatus.AppendText(ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
        //            return FileName;
        //        }

        //        if (keyword != null)
        //        {
        //            try
        //            {
        //                keyword.DeleteParameter(1);
        //                keyword.InsertNewParameterFileName(NewFileName, 1);
        //                //FileName = keyword.GetParameter(1).ToFileNamePath();
        //            }
        //            catch (Exception ex)
        //            {
        //                richTextBoxStatus.AppendText(ex.Message + (ex.InnerException != null ? " Inner: " + ex.InnerException.Message : ""));
        //                return FileName;
        //            }
        //        }
        //    }

        //    return FileName;
        //}

        //private void button1_Click(object sender, EventArgs e)
        //{
        //    using (CSSPWebToolsDBEntities db = new CSSPWebToolsDBEntities())
        //    {
        //        List<TVFile> tvFileList = (from c in db.TVFiles
        //                                   where c.FileType == (int)FileTypeEnum.M21FM
        //                                   || c.FileType == (int)FileTypeEnum.M3FM
        //                                   orderby c.TVFileTVItemID
        //                                   select c).ToList();

        //        foreach (TVFile tvFile in tvFileList)
        //        {
        //            string FileName = tvFile.ServerFilePath + tvFile.ServerFileName;
        //            label1.Text = "(" + tvFile.TVFileTVItemID.ToString() + ") --- " + FileName;
        //            label1.Refresh();
        //            Application.DoEvents();

        //            PFSFile pfsFile = new PFSFile(FileName);

        //            for (int i = 2; i < 100; i++)
        //            {
        //                string BoundCondFileName = GetFileName(pfsFile, "FemEngineHD/HYDRODYNAMIC_MODULE/BOUNDARY_CONDITIONS/CODE_" + i.ToString(), "file_name");
        //                if (string.IsNullOrWhiteSpace(BoundCondFileName))
        //                    break;

        //                FileInfo fi = new FileInfo(BoundCondFileName);

        //                string BoundCondNewFileName = SetFileName(pfsFile, "FemEngineHD/HYDRODYNAMIC_MODULE/BOUNDARY_CONDITIONS/CODE_" + i.ToString(), "file_name", @".\" + fi.Name);
        //                BoundCondFileName = GetFileName(pfsFile, "FemEngineHD/HYDRODYNAMIC_MODULE/BOUNDARY_CONDITIONS/CODE_" + i.ToString(), "file_name");
        //                fi = new FileInfo(BoundCondFileName);

        //                if (!fi.Exists)
        //                {
        //                    // file does not exist
        //                    int seflij = 234;
        //                }
        //            }

        //            pfsFile.Write(FileName);

        //            pfsFile.Close();
        //            FixPFSFileSystemPart(FileName);
        //        }
        //    }
        //}
        //private void FixPFSFileSystemPart(string FullFileName)
        //{
        //    StreamReader sr = new StreamReader(FullFileName, Encoding.GetEncoding("iso-8859-1"));

        //    FileInfo fi = new FileInfo(FullFileName);
        //    string FileText = sr.ReadToEnd();
        //    sr.Close();

        //    if (FileText.IndexOf("[SYSTEM]") > 0)
        //    {
        //        FileText = FileText.Substring(0, FileText.IndexOf("[SYSTEM]"));
        //    }
        //    StringBuilder sb = new StringBuilder(FileText);
        //    sb.AppendLine(@"[SYSTEM]");
        //    sb.AppendLine(@"   ResultRootFolder = ||");
        //    sb.AppendLine(@"   UseCustomResultFolder = true");
        //    sb.AppendLine(@"   CustomResultFolder = |.\|");
        //    sb.AppendLine(@"EndSect  // SYSTEM");

        //    StreamWriter sw = new StreamWriter(FullFileName, false, Encoding.GetEncoding("iso-8859-1"));
        //    sw.Write(sb.ToString());
        //    sw.Close();
        //}
    }

    public class YearMinMaxDate
    {
        public int Year { get; set; }
        public DateTime MinDateTime { get; set; }
        public DateTime MaxDateTime { get; set; }
    }


}
