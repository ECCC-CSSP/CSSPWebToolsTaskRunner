﻿namespace CSSPWebToolsTaskRunner
{
    partial class CSSPWebToolsTaskRunner
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(CSSPWebToolsTaskRunner));
            this.splitContainer1 = new System.Windows.Forms.SplitContainer();
            this.button1 = new System.Windows.Forms.Button();
            this.lblLastAppTaskCheckDate = new System.Windows.Forms.Label();
            this.lblLastAppTaskCheck = new System.Windows.Forms.Label();
            this.richTextBoxStatus = new System.Windows.Forms.RichTextBox();
            this.timerCheckTask = new System.Windows.Forms.Timer(this.components);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).BeginInit();
            this.splitContainer1.Panel1.SuspendLayout();
            this.splitContainer1.Panel2.SuspendLayout();
            this.splitContainer1.SuspendLayout();
            this.SuspendLayout();
            // 
            // splitContainer1
            // 
            this.splitContainer1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer1.Location = new System.Drawing.Point(0, 0);
            this.splitContainer1.Name = "splitContainer1";
            this.splitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal;
            // 
            // splitContainer1.Panel1
            // 
            this.splitContainer1.Panel1.Controls.Add(this.button1);
            this.splitContainer1.Panel1.Controls.Add(this.lblLastAppTaskCheckDate);
            this.splitContainer1.Panel1.Controls.Add(this.lblLastAppTaskCheck);
            // 
            // splitContainer1.Panel2
            // 
            this.splitContainer1.Panel2.Controls.Add(this.richTextBoxStatus);
            this.splitContainer1.Size = new System.Drawing.Size(1151, 705);
            this.splitContainer1.SplitterDistance = 71;
            this.splitContainer1.TabIndex = 0;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(531, 23);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 2;
            this.button1.Text = "button1";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // lblLastAppTaskCheckDate
            // 
            this.lblLastAppTaskCheckDate.AutoSize = true;
            this.lblLastAppTaskCheckDate.Location = new System.Drawing.Point(132, 13);
            this.lblLastAppTaskCheckDate.Name = "lblLastAppTaskCheckDate";
            this.lblLastAppTaskCheckDate.Size = new System.Drawing.Size(41, 13);
            this.lblLastAppTaskCheckDate.TabIndex = 1;
            this.lblLastAppTaskCheckDate.Text = "[empty]";
            // 
            // lblLastAppTaskCheck
            // 
            this.lblLastAppTaskCheck.AutoSize = true;
            this.lblLastAppTaskCheck.Location = new System.Drawing.Point(13, 13);
            this.lblLastAppTaskCheck.Name = "lblLastAppTaskCheck";
            this.lblLastAppTaskCheck.Size = new System.Drawing.Size(113, 13);
            this.lblLastAppTaskCheck.TabIndex = 0;
            this.lblLastAppTaskCheck.Text = "Last AppTask Check: ";
            // 
            // richTextBoxStatus
            // 
            this.richTextBoxStatus.Dock = System.Windows.Forms.DockStyle.Fill;
            this.richTextBoxStatus.Location = new System.Drawing.Point(0, 0);
            this.richTextBoxStatus.Name = "richTextBoxStatus";
            this.richTextBoxStatus.Size = new System.Drawing.Size(1151, 630);
            this.richTextBoxStatus.TabIndex = 0;
            this.richTextBoxStatus.Text = "";
            // 
            // timerCheckTask
            // 
            this.timerCheckTask.Interval = 1000;
            this.timerCheckTask.Tick += new System.EventHandler(this.timerCheckTask_Tick);
            // 
            // CSSPWebToolsTaskRunner
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1151, 705);
            this.Controls.Add(this.splitContainer1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "CSSPWebToolsTaskRunner";
            this.Text = "CSSP Web Tools Task Runner";
            this.splitContainer1.Panel1.ResumeLayout(false);
            this.splitContainer1.Panel1.PerformLayout();
            this.splitContainer1.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).EndInit();
            this.splitContainer1.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.SplitContainer splitContainer1;
        private System.Windows.Forms.RichTextBox richTextBoxStatus;
        private System.Windows.Forms.Timer timerCheckTask;
        private System.Windows.Forms.Label lblLastAppTaskCheckDate;
        private System.Windows.Forms.Label lblLastAppTaskCheck;
        private System.Windows.Forms.Button button1;
    }
}

