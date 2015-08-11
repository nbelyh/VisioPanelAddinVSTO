#region Copyright ©2005, Cristi Potlog - All Rights Reserved
/* ------------------------------------------------------------------- *
*                            Cristi Potlog                             *
*                  Copyright ©2005 - All Rights reserved               *
*                                                                      *
* THIS SOURCE CODE IS PROVIDED "AS IS" WITH NO WARRANTIES OF ANY KIND, *
* EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE        *
* WARRANTIES OF DESIGN, MERCHANTIBILITY AND FITNESS FOR A PARTICULAR   *
* PURPOSE, NONINFRINGEMENT, OR ARISING FROM A COURSE OF DEALING,       *
* USAGE OR TRADE PRACTICE.                                             *
*                                                                      *
* THIS COPYRIGHT NOTICE MAY NOT BE REMOVED FROM THIS FILE.             *
* ------------------------------------------------------------------- */
#endregion Copyright ©2005, Cristi Potlog - All Rights Reserved

#region References

using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;
using System.Linq;

#endregion

namespace PanelAddinWizard
{
	/// <summary>
	/// Summary description for WizardForm.
	/// </summary>
	public class WizardForm : Form
	{
        /// <summary>
        /// Creates a new instance of the <see cref="WizardForm"/> class.
        /// </summary>
        public WizardForm(bool wixInstalled)
        {
            // required for designer support
            InitializeComponent();

            checkAddinProject.Checked = true;

            checkWixSetup.Enabled = wixInstalled;
            checkWixSetup.Checked = wixInstalled;

            checkAddVisioFiles.Checked = false;
            radioCreateNewVisioFiles.Checked = true;
            checkCopyVisioFiles.Checked = true;

            UpdateButtons(null, null);
        }

        public bool TaskPane { get { return checkSupportTaskPane.Checked; } }
        public bool Ribbon { get { return checkSupportRibbon.Checked; } }
        public bool CommandBars { get { return checkSupportCommandBars.Checked; } }
        public bool AddinEnabled { get { return checkAddinProject.Checked; } }

        public bool WixSetup { get { return checkWixSetup.Checked; } }

        private Label radioUseVisioFilesDescription;
        private CheckBox checkAddVisioFiles;
        private Label radioCreateNewVisioFilesDescription;
        private TextBox textBoxVisioFilesPath;
        private CheckBox checkCopyVisioFiles;
        private Button buttonBrowseVisioFiles;
        private RadioButton radioCreateNewVisioFiles;
        private RadioButton radioUseVisioFiles;
        private LinkLabel checkAddinProjectDescription;
        private CheckBox checkAddinProject;
        string[] VisioFilePaths;

        public WixSetupOptions GetSetupOptions()
        {
            return new WixSetupOptions
            {
                Enabled = checkAddVisioFiles.Checked,
                CreateNew = radioCreateNewVisioFiles.Checked,
                Duplicate = checkCopyVisioFiles.Checked,
                Paths = this.VisioFilePaths,
            };
        }

        public Image HeaderImage
	    {
	        get { return addinWizard.HeaderImage; }
            set { addinWizard.HeaderImage = value; }
	    }

		#region Designer generated code

        private WizardPage pageAddin;
        private Wizard addinWizard;
        private WizardPage pageSetup;
        private Label checkSupportTaskPaneDescription;
        private CheckBox checkSupportTaskPane;
        private Label checkSupportCommandBarsDescription;
        private Label checkSupportRibbonDescription;
        private CheckBox checkSupportCommandBars;
        private CheckBox checkSupportRibbon;
        private LinkLabel checkWixSetupDescription;
        private CheckBox checkWixSetup;
		
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(WizardForm));
            this.addinWizard = new PanelAddinWizard.Wizard();
            this.pageSetup = new PanelAddinWizard.WizardPage();
            this.radioUseVisioFilesDescription = new System.Windows.Forms.Label();
            this.checkAddVisioFiles = new System.Windows.Forms.CheckBox();
            this.radioCreateNewVisioFilesDescription = new System.Windows.Forms.Label();
            this.checkWixSetupDescription = new System.Windows.Forms.LinkLabel();
            this.textBoxVisioFilesPath = new System.Windows.Forms.TextBox();
            this.checkWixSetup = new System.Windows.Forms.CheckBox();
            this.checkCopyVisioFiles = new System.Windows.Forms.CheckBox();
            this.radioUseVisioFiles = new System.Windows.Forms.RadioButton();
            this.buttonBrowseVisioFiles = new System.Windows.Forms.Button();
            this.radioCreateNewVisioFiles = new System.Windows.Forms.RadioButton();
            this.pageAddin = new PanelAddinWizard.WizardPage();
            this.checkAddinProjectDescription = new System.Windows.Forms.LinkLabel();
            this.checkAddinProject = new System.Windows.Forms.CheckBox();
            this.checkSupportTaskPaneDescription = new System.Windows.Forms.Label();
            this.checkSupportTaskPane = new System.Windows.Forms.CheckBox();
            this.checkSupportCommandBarsDescription = new System.Windows.Forms.Label();
            this.checkSupportRibbonDescription = new System.Windows.Forms.Label();
            this.checkSupportCommandBars = new System.Windows.Forms.CheckBox();
            this.checkSupportRibbon = new System.Windows.Forms.CheckBox();
            this.addinWizard.SuspendLayout();
            this.pageSetup.SuspendLayout();
            this.pageAddin.SuspendLayout();
            this.SuspendLayout();
            // 
            // addinWizard
            // 
            this.addinWizard.Controls.Add(this.pageSetup);
            this.addinWizard.Controls.Add(this.pageAddin);
            this.addinWizard.HelpVisible = true;
            this.addinWizard.Location = new System.Drawing.Point(0, 0);
            this.addinWizard.Name = "addinWizard";
            this.addinWizard.Pages.AddRange(new PanelAddinWizard.WizardPage[] {
            this.pageAddin,
            this.pageSetup});
            this.addinWizard.Size = new System.Drawing.Size(557, 447);
            this.addinWizard.TabIndex = 0;
            this.addinWizard.BeforeSwitchPages += new PanelAddinWizard.Wizard.BeforeSwitchPagesEventHandler(this.addinWizard_BeforeSwitchPages);
            this.addinWizard.AfterSwitchPages += new PanelAddinWizard.Wizard.AfterSwitchPagesEventHandler(this.addinWizard_AfterSwitchPages);
            this.addinWizard.Cancel += new System.ComponentModel.CancelEventHandler(this.addinWizard_Cancel);
            this.addinWizard.Help += new System.EventHandler(this.addinWizard_Help);
            // 
            // pageSetup
            // 
            this.pageSetup.Controls.Add(this.radioUseVisioFilesDescription);
            this.pageSetup.Controls.Add(this.checkAddVisioFiles);
            this.pageSetup.Controls.Add(this.radioCreateNewVisioFilesDescription);
            this.pageSetup.Controls.Add(this.checkWixSetupDescription);
            this.pageSetup.Controls.Add(this.textBoxVisioFilesPath);
            this.pageSetup.Controls.Add(this.checkWixSetup);
            this.pageSetup.Controls.Add(this.checkCopyVisioFiles);
            this.pageSetup.Controls.Add(this.radioUseVisioFiles);
            this.pageSetup.Controls.Add(this.buttonBrowseVisioFiles);
            this.pageSetup.Controls.Add(this.radioCreateNewVisioFiles);
            this.pageSetup.Description = "Please select deployment options";
            this.pageSetup.Location = new System.Drawing.Point(0, 0);
            this.pageSetup.Name = "pageSetup";
            this.pageSetup.Size = new System.Drawing.Size(557, 399);
            this.pageSetup.TabIndex = 0;
            this.pageSetup.Title = "Setup project";
            // 
            // radioUseVisioFilesDescription
            // 
            this.radioUseVisioFilesDescription.ForeColor = System.Drawing.SystemColors.GrayText;
            this.radioUseVisioFilesDescription.Location = new System.Drawing.Point(73, 266);
            this.radioUseVisioFilesDescription.Name = "radioUseVisioFilesDescription";
            this.radioUseVisioFilesDescription.Size = new System.Drawing.Size(380, 30);
            this.radioUseVisioFilesDescription.TabIndex = 6;
            this.radioUseVisioFilesDescription.Text = "Choose this option if you already have Visio template or stencil file(s) and want" +
    " to include it in the project.";
            // 
            // checkAddVisioFiles
            // 
            this.checkAddVisioFiles.AutoSize = true;
            this.checkAddVisioFiles.Location = new System.Drawing.Point(32, 161);
            this.checkAddVisioFiles.Name = "checkAddVisioFiles";
            this.checkAddVisioFiles.Size = new System.Drawing.Size(155, 17);
            this.checkAddVisioFiles.TabIndex = 2;
            this.checkAddVisioFiles.Text = "Include custom Visio files(s)";
            this.checkAddVisioFiles.UseVisualStyleBackColor = true;
            this.checkAddVisioFiles.CheckedChanged += new System.EventHandler(this.UpdateButtons);
            // 
            // radioCreateNewVisioFilesDescription
            // 
            this.radioCreateNewVisioFilesDescription.ForeColor = System.Drawing.SystemColors.GrayText;
            this.radioCreateNewVisioFilesDescription.Location = new System.Drawing.Point(70, 212);
            this.radioCreateNewVisioFilesDescription.Name = "radioCreateNewVisioFilesDescription";
            this.radioCreateNewVisioFilesDescription.Size = new System.Drawing.Size(383, 28);
            this.radioCreateNewVisioFilesDescription.TabIndex = 4;
            this.radioCreateNewVisioFilesDescription.Text = "Creates a sample template and includes it in the project. You can use it as a sta" +
    "rting point for your Visio project.";
            // 
            // checkWixSetupDescription
            // 
            this.checkWixSetupDescription.ForeColor = System.Drawing.SystemColors.GrayText;
            this.checkWixSetupDescription.LinkArea = new System.Windows.Forms.LinkArea(196, 21);
            this.checkWixSetupDescription.LinkColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(192)))));
            this.checkWixSetupDescription.Location = new System.Drawing.Point(32, 100);
            this.checkWixSetupDescription.Name = "checkWixSetupDescription";
            this.checkWixSetupDescription.Size = new System.Drawing.Size(513, 46);
            this.checkWixSetupDescription.TabIndex = 1;
            this.checkWixSetupDescription.TabStop = true;
            this.checkWixSetupDescription.Text = resources.GetString("checkWixSetupDescription.Text");
            this.checkWixSetupDescription.UseCompatibleTextRendering = true;
            this.checkWixSetupDescription.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.checkWixSetupDescription_LinkClicked);
            // 
            // textBoxVisioFilesPath
            // 
            this.textBoxVisioFilesPath.Location = new System.Drawing.Point(73, 299);
            this.textBoxVisioFilesPath.Name = "textBoxVisioFilesPath";
            this.textBoxVisioFilesPath.ReadOnly = true;
            this.textBoxVisioFilesPath.Size = new System.Drawing.Size(380, 20);
            this.textBoxVisioFilesPath.TabIndex = 7;
            // 
            // checkWixSetup
            // 
            this.checkWixSetup.AutoSize = true;
            this.checkWixSetup.Location = new System.Drawing.Point(16, 80);
            this.checkWixSetup.Name = "checkWixSetup";
            this.checkWixSetup.Size = new System.Drawing.Size(121, 17);
            this.checkWixSetup.TabIndex = 0;
            this.checkWixSetup.Text = "Create setup project";
            this.checkWixSetup.UseVisualStyleBackColor = true;
            this.checkWixSetup.Click += new System.EventHandler(this.UpdateButtons);
            // 
            // checkCopyVisioFiles
            // 
            this.checkCopyVisioFiles.AutoSize = true;
            this.checkCopyVisioFiles.Location = new System.Drawing.Point(73, 325);
            this.checkCopyVisioFiles.Name = "checkCopyVisioFiles";
            this.checkCopyVisioFiles.Size = new System.Drawing.Size(185, 17);
            this.checkCopyVisioFiles.TabIndex = 9;
            this.checkCopyVisioFiles.Text = "Copy file(s) to the project directory";
            this.checkCopyVisioFiles.UseVisualStyleBackColor = true;
            // 
            // radioUseVisioFiles
            // 
            this.radioUseVisioFiles.AutoSize = true;
            this.radioUseVisioFiles.Location = new System.Drawing.Point(51, 246);
            this.radioUseVisioFiles.Name = "radioUseVisioFiles";
            this.radioUseVisioFiles.Size = new System.Drawing.Size(224, 17);
            this.radioUseVisioFiles.TabIndex = 5;
            this.radioUseVisioFiles.TabStop = true;
            this.radioUseVisioFiles.Text = "Use already existing template/stencil file(s)";
            this.radioUseVisioFiles.UseVisualStyleBackColor = true;
            // 
            // buttonBrowseVisioFiles
            // 
            this.buttonBrowseVisioFiles.Location = new System.Drawing.Point(459, 297);
            this.buttonBrowseVisioFiles.Name = "buttonBrowseVisioFiles";
            this.buttonBrowseVisioFiles.Size = new System.Drawing.Size(75, 23);
            this.buttonBrowseVisioFiles.TabIndex = 8;
            this.buttonBrowseVisioFiles.Text = "Browse...";
            this.buttonBrowseVisioFiles.UseVisualStyleBackColor = true;
            this.buttonBrowseVisioFiles.Click += new System.EventHandler(this.buttonBrowse_Click);
            // 
            // radioCreateNewVisioFiles
            // 
            this.radioCreateNewVisioFiles.AutoSize = true;
            this.radioCreateNewVisioFiles.Location = new System.Drawing.Point(51, 188);
            this.radioCreateNewVisioFiles.Name = "radioCreateNewVisioFiles";
            this.radioCreateNewVisioFiles.Size = new System.Drawing.Size(243, 17);
            this.radioCreateNewVisioFiles.TabIndex = 3;
            this.radioCreateNewVisioFiles.TabStop = true;
            this.radioCreateNewVisioFiles.Text = "Create a new (sample) template and stencil file";
            this.radioCreateNewVisioFiles.UseVisualStyleBackColor = true;
            this.radioCreateNewVisioFiles.CheckedChanged += new System.EventHandler(this.UpdateButtons);
            // 
            // pageAddin
            // 
            this.pageAddin.Controls.Add(this.checkAddinProjectDescription);
            this.pageAddin.Controls.Add(this.checkAddinProject);
            this.pageAddin.Controls.Add(this.checkSupportTaskPaneDescription);
            this.pageAddin.Controls.Add(this.checkSupportTaskPane);
            this.pageAddin.Controls.Add(this.checkSupportCommandBarsDescription);
            this.pageAddin.Controls.Add(this.checkSupportRibbonDescription);
            this.pageAddin.Controls.Add(this.checkSupportCommandBars);
            this.pageAddin.Controls.Add(this.checkSupportRibbon);
            this.pageAddin.Description = "Please select add-in features";
            this.pageAddin.Location = new System.Drawing.Point(0, 0);
            this.pageAddin.Name = "pageAddin";
            this.pageAddin.Size = new System.Drawing.Size(557, 399);
            this.pageAddin.TabIndex = 11;
            this.pageAddin.Title = "Add-in options";
            // 
            // checkAddinProjectDescription
            // 
            this.checkAddinProjectDescription.ForeColor = System.Drawing.SystemColors.GrayText;
            this.checkAddinProjectDescription.LinkArea = new System.Windows.Forms.LinkArea(187, 8);
            this.checkAddinProjectDescription.LinkColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(192)))));
            this.checkAddinProjectDescription.Location = new System.Drawing.Point(32, 100);
            this.checkAddinProjectDescription.Name = "checkAddinProjectDescription";
            this.checkAddinProjectDescription.Size = new System.Drawing.Size(513, 46);
            this.checkAddinProjectDescription.TabIndex = 16;
            this.checkAddinProjectDescription.TabStop = true;
            this.checkAddinProjectDescription.Text = resources.GetString("checkAddinProjectDescription.Text");
            this.checkAddinProjectDescription.UseCompatibleTextRendering = true;
            // 
            // checkAddinProject
            // 
            this.checkAddinProject.AutoSize = true;
            this.checkAddinProject.Location = new System.Drawing.Point(16, 80);
            this.checkAddinProject.Name = "checkAddinProject";
            this.checkAddinProject.Size = new System.Drawing.Size(147, 17);
            this.checkAddinProject.TabIndex = 15;
            this.checkAddinProject.Text = "Create Visio Addin project";
            this.checkAddinProject.UseVisualStyleBackColor = true;
            this.checkAddinProject.CheckedChanged += new System.EventHandler(this.UpdateButtons);
            // 
            // checkSupportTaskPaneDescription
            // 
            this.checkSupportTaskPaneDescription.ForeColor = System.Drawing.SystemColors.GrayText;
            this.checkSupportTaskPaneDescription.Location = new System.Drawing.Point(48, 184);
            this.checkSupportTaskPaneDescription.Name = "checkSupportTaskPaneDescription";
            this.checkSupportTaskPaneDescription.Size = new System.Drawing.Size(380, 24);
            this.checkSupportTaskPaneDescription.TabIndex = 12;
            this.checkSupportTaskPaneDescription.Text = "Adds a docking panel which can be controlelled with a toggle button";
            // 
            // checkSupportTaskPane
            // 
            this.checkSupportTaskPane.AutoSize = true;
            this.checkSupportTaskPane.Checked = true;
            this.checkSupportTaskPane.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkSupportTaskPane.Location = new System.Drawing.Point(32, 160);
            this.checkSupportTaskPane.Name = "checkSupportTaskPane";
            this.checkSupportTaskPane.Size = new System.Drawing.Size(118, 17);
            this.checkSupportTaskPane.TabIndex = 9;
            this.checkSupportTaskPane.Text = "Support Task Pane";
            this.checkSupportTaskPane.UseVisualStyleBackColor = true;
            // 
            // checkSupportCommandBarsDescription
            // 
            this.checkSupportCommandBarsDescription.ForeColor = System.Drawing.SystemColors.GrayText;
            this.checkSupportCommandBarsDescription.Location = new System.Drawing.Point(48, 280);
            this.checkSupportCommandBarsDescription.Name = "checkSupportCommandBarsDescription";
            this.checkSupportCommandBarsDescription.Size = new System.Drawing.Size(380, 24);
            this.checkSupportCommandBarsDescription.TabIndex = 14;
            this.checkSupportCommandBarsDescription.Text = "Add a toolbar with custom images (check if you need to support old Visio)";
            // 
            // checkSupportRibbonDescription
            // 
            this.checkSupportRibbonDescription.ForeColor = System.Drawing.SystemColors.GrayText;
            this.checkSupportRibbonDescription.Location = new System.Drawing.Point(48, 232);
            this.checkSupportRibbonDescription.Name = "checkSupportRibbonDescription";
            this.checkSupportRibbonDescription.Size = new System.Drawing.Size(380, 24);
            this.checkSupportRibbonDescription.TabIndex = 13;
            this.checkSupportRibbonDescription.Text = "Add a ribbon with buttons and custom images and state";
            // 
            // checkSupportCommandBars
            // 
            this.checkSupportCommandBars.AutoSize = true;
            this.checkSupportCommandBars.Location = new System.Drawing.Point(32, 256);
            this.checkSupportCommandBars.Name = "checkSupportCommandBars";
            this.checkSupportCommandBars.Size = new System.Drawing.Size(247, 17);
            this.checkSupportCommandBars.TabIndex = 11;
            this.checkSupportCommandBars.Text = "Support Command Bars (Visio 2007 and below)";
            this.checkSupportCommandBars.UseVisualStyleBackColor = true;
            // 
            // checkSupportRibbon
            // 
            this.checkSupportRibbon.AutoSize = true;
            this.checkSupportRibbon.Checked = true;
            this.checkSupportRibbon.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkSupportRibbon.Location = new System.Drawing.Point(32, 208);
            this.checkSupportRibbon.Name = "checkSupportRibbon";
            this.checkSupportRibbon.Size = new System.Drawing.Size(212, 17);
            this.checkSupportRibbon.TabIndex = 10;
            this.checkSupportRibbon.Text = "Support Ribbon (Visio 2010 and above)";
            this.checkSupportRibbon.UseVisualStyleBackColor = true;
            // 
            // WizardForm
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            this.ClientSize = new System.Drawing.Size(557, 447);
            this.Controls.Add(this.addinWizard);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "WizardForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Create New Visio Project";
            this.addinWizard.ResumeLayout(false);
            this.pageSetup.ResumeLayout(false);
            this.pageSetup.PerformLayout();
            this.pageAddin.ResumeLayout(false);
            this.pageAddin.PerformLayout();
            this.ResumeLayout(false);

        }
		#endregion

	    private void addinWizard_AfterSwitchPages(object sender, Wizard.AfterSwitchPagesEventArgs e)
		{
            UpdateButtons(null, null);
		}

		private void addinWizard_BeforeSwitchPages(object sender, Wizard.BeforeSwitchPagesEventArgs e)
		{
		}

		private void addinWizard_Cancel(object sender, CancelEventArgs e)
		{
		}

        private void addinWizard_Help(object sender, EventArgs e)
        {
        }

        private void UpdateButtons(object sender, EventArgs e)
        {
            checkAddVisioFiles.Enabled = checkWixSetup.Checked;
            radioCreateNewVisioFiles.Enabled = checkWixSetup.Checked && checkAddVisioFiles.Checked;
            radioCreateNewVisioFilesDescription.ForeColor = radioCreateNewVisioFiles.Enabled ? SystemColors.GrayText : SystemColors.ControlDark;

            radioUseVisioFiles.Enabled = checkWixSetup.Checked && checkAddVisioFiles.Checked;
            radioUseVisioFilesDescription.ForeColor = radioUseVisioFiles.Enabled ? SystemColors.GrayText : SystemColors.ControlDark;

            textBoxVisioFilesPath.Enabled = checkWixSetup.Checked && checkAddVisioFiles.Checked && radioUseVisioFiles.Checked;

            buttonBrowseVisioFiles.Enabled = textBoxVisioFilesPath.Enabled;
            checkCopyVisioFiles.Enabled = textBoxVisioFilesPath.Enabled;

            textBoxVisioFilesPath.Text = VisioFilePaths == null ? ""
                : string.Join(" ", VisioFilePaths.Select(p => string.Format(@"""{0}""", System.IO.Path.GetFileName(p))));

            checkSupportTaskPane.Enabled = checkAddinProject.Checked;
            checkSupportTaskPaneDescription.ForeColor = checkSupportTaskPane.Enabled ? SystemColors.GrayText : SystemColors.ControlDark;

            checkSupportRibbon.Enabled = checkAddinProject.Checked;
            checkSupportRibbonDescription.ForeColor = checkSupportRibbon.Enabled ? SystemColors.GrayText : SystemColors.ControlDark;

            checkSupportCommandBars.Enabled = checkAddinProject.Checked;
            checkSupportCommandBarsDescription.ForeColor = checkSupportCommandBars.Enabled ? SystemColors.GrayText : SystemColors.ControlDark;

            addinWizard.NextEnabled = !(checkWixSetup.Checked && checkAddVisioFiles.Checked && radioUseVisioFiles.Checked && textBoxVisioFilesPath.Text.Length == 0);
        }
        
        private void checkWixSetupDescription_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Process.Start(new ProcessStartInfo("http://wixtoolset.org/"));
        }

        private void buttonBrowse_Click(object sender, EventArgs e)
        {
            var dlg = new OpenFileDialog
            {
                Filter =
                    "Visio templates and stencils|*.vst;*.vtx;*.vst?;*.vss;*.vsx;*.vss?|" +
                    "Visio templates|*.vst;*.vtx;*.vst?|" +
                    "Visio stencils|*.vss;*.vsx;*.vss?|" +
                    "All Files (*.*)|*.*",
                Multiselect = true
            };

            if (dlg.ShowDialog() == DialogResult.OK)
            {
                VisioFilePaths = dlg.FileNames;
                UpdateButtons(null, null);
            }
        }
    }
}
