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

            checkWixSetup.Enabled = wixInstalled;
            radioCreateNewStencil.Checked = true;
            radioCreateNewTemplate.Checked = true;

            UpdateButtons(null, null);
        }

        public bool TaskPane { get { return checkSupportTaskPane.Checked; } }
        public bool Ribbon { get { return checkSupportRibbon.Checked; } }
        public bool CommandBars { get { return checkSupportCommandBars.Checked; } }

        public bool WixSetup { get { return checkWixSetup.Checked; } }

        public bool AddVisioTemplates { get { return checkAddVisioTemplates.Checked; } }
        public bool CreateNewTemplate { get { return radioCreateNewTemplate.Checked; } }
        public bool UseTemplate { get { return radioUseTemplate.Checked; } }
        public string[] TemplatePaths { get; private set; }

        public bool AddVisioStencils { get { return checkAddVisioStencils.Checked; } }
        public bool CreateNewStencil { get { return radioCreateNewStencil.Checked; } }
        public bool UseStencil { get { return radioUseStencil.Checked; } }
        public string[] StencilPaths { get; private set; }

	    public Image HeaderImage
	    {
	        get { return addinWizard.HeaderImage; }
            set { addinWizard.HeaderImage = value; }
	    }

		#region Designer generated code

        private WizardPage pageOptions;
        private Wizard addinWizard;
        private WizardPage pageCheck;
        private Label labelSupportTaskPane;
        private CheckBox checkSupportTaskPane;
        private Label labelSupportCommandBars;
        private Label labelSupportRibbon;
        private CheckBox checkSupportCommandBars;
        private CheckBox checkSupportRibbon;
        private CheckBox checkAddVisioStencils;
        private CheckBox checkAddVisioTemplates;
        private LinkLabel checkWixSetupDescription;
        private TextBox textBoxStencilPath;
        private RadioButton radioUseStencil;
        private RadioButton radioCreateNewStencil;
        private TextBox textBoxTemplatePath;
        private RadioButton radioUseTemplate;
        private RadioButton radioCreateNewTemplate;
        private Button buttonBrowseStencil;
        private Button buttonBrowseTemplate;
        private Panel panelStencil;
        private Panel panelTemplate;
        private CheckBox checkWixSetup;
		
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(WizardForm));
            this.addinWizard = new PanelAddinWizard.Wizard();
            this.pageOptions = new PanelAddinWizard.WizardPage();
            this.labelSupportTaskPane = new System.Windows.Forms.Label();
            this.checkSupportTaskPane = new System.Windows.Forms.CheckBox();
            this.labelSupportCommandBars = new System.Windows.Forms.Label();
            this.labelSupportRibbon = new System.Windows.Forms.Label();
            this.checkSupportCommandBars = new System.Windows.Forms.CheckBox();
            this.checkSupportRibbon = new System.Windows.Forms.CheckBox();
            this.pageCheck = new PanelAddinWizard.WizardPage();
            this.panelStencil = new System.Windows.Forms.Panel();
            this.checkAddVisioStencils = new System.Windows.Forms.CheckBox();
            this.radioCreateNewStencil = new System.Windows.Forms.RadioButton();
            this.buttonBrowseStencil = new System.Windows.Forms.Button();
            this.radioUseStencil = new System.Windows.Forms.RadioButton();
            this.textBoxStencilPath = new System.Windows.Forms.TextBox();
            this.panelTemplate = new System.Windows.Forms.Panel();
            this.radioCreateNewTemplate = new System.Windows.Forms.RadioButton();
            this.radioUseTemplate = new System.Windows.Forms.RadioButton();
            this.buttonBrowseTemplate = new System.Windows.Forms.Button();
            this.textBoxTemplatePath = new System.Windows.Forms.TextBox();
            this.checkAddVisioTemplates = new System.Windows.Forms.CheckBox();
            this.checkWixSetupDescription = new System.Windows.Forms.LinkLabel();
            this.checkWixSetup = new System.Windows.Forms.CheckBox();
            this.addinWizard.SuspendLayout();
            this.pageOptions.SuspendLayout();
            this.pageCheck.SuspendLayout();
            this.panelStencil.SuspendLayout();
            this.panelTemplate.SuspendLayout();
            this.SuspendLayout();
            // 
            // addinWizard
            // 
            this.addinWizard.Controls.Add(this.pageCheck);
            this.addinWizard.Controls.Add(this.pageOptions);
            this.addinWizard.HelpVisible = true;
            this.addinWizard.Location = new System.Drawing.Point(0, 0);
            this.addinWizard.Name = "addinWizard";
            this.addinWizard.Pages.AddRange(new PanelAddinWizard.WizardPage[] {
            this.pageOptions,
            this.pageCheck});
            this.addinWizard.Size = new System.Drawing.Size(569, 462);
            this.addinWizard.TabIndex = 0;
            this.addinWizard.BeforeSwitchPages += new PanelAddinWizard.Wizard.BeforeSwitchPagesEventHandler(this.addinWizard_BeforeSwitchPages);
            this.addinWizard.AfterSwitchPages += new PanelAddinWizard.Wizard.AfterSwitchPagesEventHandler(this.addinWizard_AfterSwitchPages);
            this.addinWizard.Cancel += new System.ComponentModel.CancelEventHandler(this.addinWizard_Cancel);
            this.addinWizard.Help += new System.EventHandler(this.addinWizard_Help);
            // 
            // pageOptions
            // 
            this.pageOptions.Controls.Add(this.labelSupportTaskPane);
            this.pageOptions.Controls.Add(this.checkSupportTaskPane);
            this.pageOptions.Controls.Add(this.labelSupportCommandBars);
            this.pageOptions.Controls.Add(this.labelSupportRibbon);
            this.pageOptions.Controls.Add(this.checkSupportCommandBars);
            this.pageOptions.Controls.Add(this.checkSupportRibbon);
            this.pageOptions.Description = "Please select add-in features";
            this.pageOptions.Location = new System.Drawing.Point(0, 0);
            this.pageOptions.Name = "pageOptions";
            this.pageOptions.Size = new System.Drawing.Size(569, 414);
            this.pageOptions.TabIndex = 11;
            this.pageOptions.Title = "Add-in options";
            // 
            // labelSupportTaskPane
            // 
            this.labelSupportTaskPane.ForeColor = System.Drawing.SystemColors.GrayText;
            this.labelSupportTaskPane.Location = new System.Drawing.Point(32, 104);
            this.labelSupportTaskPane.Name = "labelSupportTaskPane";
            this.labelSupportTaskPane.Size = new System.Drawing.Size(380, 24);
            this.labelSupportTaskPane.TabIndex = 12;
            this.labelSupportTaskPane.Text = "Adds a docking panel which can be controlelled with a toggle button";
            // 
            // checkSupportTaskPane
            // 
            this.checkSupportTaskPane.AutoSize = true;
            this.checkSupportTaskPane.Checked = true;
            this.checkSupportTaskPane.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkSupportTaskPane.Location = new System.Drawing.Point(16, 80);
            this.checkSupportTaskPane.Name = "checkSupportTaskPane";
            this.checkSupportTaskPane.Size = new System.Drawing.Size(118, 17);
            this.checkSupportTaskPane.TabIndex = 9;
            this.checkSupportTaskPane.Text = "Support Task Pane";
            this.checkSupportTaskPane.UseVisualStyleBackColor = true;
            // 
            // labelSupportCommandBars
            // 
            this.labelSupportCommandBars.ForeColor = System.Drawing.SystemColors.GrayText;
            this.labelSupportCommandBars.Location = new System.Drawing.Point(32, 200);
            this.labelSupportCommandBars.Name = "labelSupportCommandBars";
            this.labelSupportCommandBars.Size = new System.Drawing.Size(380, 24);
            this.labelSupportCommandBars.TabIndex = 14;
            this.labelSupportCommandBars.Text = "Add a toolbar with custom images (check if you need to support old Visio)";
            // 
            // labelSupportRibbon
            // 
            this.labelSupportRibbon.ForeColor = System.Drawing.SystemColors.GrayText;
            this.labelSupportRibbon.Location = new System.Drawing.Point(32, 152);
            this.labelSupportRibbon.Name = "labelSupportRibbon";
            this.labelSupportRibbon.Size = new System.Drawing.Size(380, 24);
            this.labelSupportRibbon.TabIndex = 13;
            this.labelSupportRibbon.Text = "Add a ribbon with buttons and custom images and state";
            // 
            // checkSupportCommandBars
            // 
            this.checkSupportCommandBars.AutoSize = true;
            this.checkSupportCommandBars.Location = new System.Drawing.Point(16, 176);
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
            this.checkSupportRibbon.Location = new System.Drawing.Point(16, 128);
            this.checkSupportRibbon.Name = "checkSupportRibbon";
            this.checkSupportRibbon.Size = new System.Drawing.Size(212, 17);
            this.checkSupportRibbon.TabIndex = 10;
            this.checkSupportRibbon.Text = "Support Ribbon (Visio 2010 and above)";
            this.checkSupportRibbon.UseVisualStyleBackColor = true;
            // 
            // pageCheck
            // 
            this.pageCheck.Controls.Add(this.panelStencil);
            this.pageCheck.Controls.Add(this.panelTemplate);
            this.pageCheck.Controls.Add(this.checkWixSetupDescription);
            this.pageCheck.Controls.Add(this.checkWixSetup);
            this.pageCheck.Description = "Please select deployment options";
            this.pageCheck.Location = new System.Drawing.Point(0, 0);
            this.pageCheck.Name = "pageCheck";
            this.pageCheck.Size = new System.Drawing.Size(569, 414);
            this.pageCheck.TabIndex = 0;
            this.pageCheck.Title = "Setup project";
            // 
            // panelStencil
            // 
            this.panelStencil.Controls.Add(this.checkAddVisioStencils);
            this.panelStencil.Controls.Add(this.radioCreateNewStencil);
            this.panelStencil.Controls.Add(this.buttonBrowseStencil);
            this.panelStencil.Controls.Add(this.radioUseStencil);
            this.panelStencil.Controls.Add(this.textBoxStencilPath);
            this.panelStencil.Location = new System.Drawing.Point(32, 276);
            this.panelStencil.Name = "panelStencil";
            this.panelStencil.Size = new System.Drawing.Size(513, 100);
            this.panelStencil.TabIndex = 34;
            // 
            // checkAddVisioStencils
            // 
            this.checkAddVisioStencils.AutoSize = true;
            this.checkAddVisioStencils.Location = new System.Drawing.Point(6, 3);
            this.checkAddVisioStencils.Name = "checkAddVisioStencils";
            this.checkAddVisioStencils.Size = new System.Drawing.Size(164, 17);
            this.checkAddVisioStencils.TabIndex = 7;
            this.checkAddVisioStencils.Text = "Include cusom Visio stencil(s)";
            this.checkAddVisioStencils.UseVisualStyleBackColor = true;
            this.checkAddVisioStencils.CheckedChanged += new System.EventHandler(this.UpdateButtons);
            this.checkAddVisioStencils.Click += new System.EventHandler(this.UpdateButtons);
            // 
            // radioCreateNewStencil
            // 
            this.radioCreateNewStencil.AutoSize = true;
            this.radioCreateNewStencil.Location = new System.Drawing.Point(30, 26);
            this.radioCreateNewStencil.Name = "radioCreateNewStencil";
            this.radioCreateNewStencil.Size = new System.Drawing.Size(163, 17);
            this.radioCreateNewStencil.TabIndex = 8;
            this.radioCreateNewStencil.TabStop = true;
            this.radioCreateNewStencil.Text = "Create a new (sample) stencil";
            this.radioCreateNewStencil.UseVisualStyleBackColor = true;
            this.radioCreateNewStencil.CheckedChanged += new System.EventHandler(this.UpdateButtons);
            // 
            // buttonBrowseStencil
            // 
            this.buttonBrowseStencil.Location = new System.Drawing.Point(368, 70);
            this.buttonBrowseStencil.Name = "buttonBrowseStencil";
            this.buttonBrowseStencil.Size = new System.Drawing.Size(75, 23);
            this.buttonBrowseStencil.TabIndex = 32;
            this.buttonBrowseStencil.Text = "Browse...";
            this.buttonBrowseStencil.UseVisualStyleBackColor = true;
            this.buttonBrowseStencil.Click += new System.EventHandler(this.buttonBrowseStencil_Click);
            // 
            // radioUseStencil
            // 
            this.radioUseStencil.AutoSize = true;
            this.radioUseStencil.Location = new System.Drawing.Point(30, 49);
            this.radioUseStencil.Name = "radioUseStencil";
            this.radioUseStencil.Size = new System.Drawing.Size(130, 17);
            this.radioUseStencil.TabIndex = 9;
            this.radioUseStencil.TabStop = true;
            this.radioUseStencil.Text = "Use an existing stencil";
            this.radioUseStencil.UseVisualStyleBackColor = true;
            this.radioUseStencil.CheckedChanged += new System.EventHandler(this.UpdateButtons);
            // 
            // textBoxStencilPath
            // 
            this.textBoxStencilPath.Location = new System.Drawing.Point(30, 72);
            this.textBoxStencilPath.Name = "textBoxStencilPath";
            this.textBoxStencilPath.ReadOnly = true;
            this.textBoxStencilPath.Size = new System.Drawing.Size(332, 20);
            this.textBoxStencilPath.TabIndex = 10;
            // 
            // panelTemplate
            // 
            this.panelTemplate.Controls.Add(this.radioCreateNewTemplate);
            this.panelTemplate.Controls.Add(this.radioUseTemplate);
            this.panelTemplate.Controls.Add(this.buttonBrowseTemplate);
            this.panelTemplate.Controls.Add(this.textBoxTemplatePath);
            this.panelTemplate.Controls.Add(this.checkAddVisioTemplates);
            this.panelTemplate.Location = new System.Drawing.Point(32, 161);
            this.panelTemplate.Name = "panelTemplate";
            this.panelTemplate.Size = new System.Drawing.Size(513, 104);
            this.panelTemplate.TabIndex = 33;
            // 
            // radioCreateNewTemplate
            // 
            this.radioCreateNewTemplate.AutoSize = true;
            this.radioCreateNewTemplate.Location = new System.Drawing.Point(30, 28);
            this.radioCreateNewTemplate.Name = "radioCreateNewTemplate";
            this.radioCreateNewTemplate.Size = new System.Drawing.Size(173, 17);
            this.radioCreateNewTemplate.TabIndex = 3;
            this.radioCreateNewTemplate.TabStop = true;
            this.radioCreateNewTemplate.Text = "Create a new (sample) template";
            this.radioCreateNewTemplate.UseVisualStyleBackColor = true;
            this.radioCreateNewTemplate.CheckedChanged += new System.EventHandler(this.UpdateButtons);
            // 
            // radioUseTemplate
            // 
            this.radioUseTemplate.AutoSize = true;
            this.radioUseTemplate.Location = new System.Drawing.Point(30, 51);
            this.radioUseTemplate.Name = "radioUseTemplate";
            this.radioUseTemplate.Size = new System.Drawing.Size(140, 17);
            this.radioUseTemplate.TabIndex = 4;
            this.radioUseTemplate.TabStop = true;
            this.radioUseTemplate.Text = "Use an existing template";
            this.radioUseTemplate.UseVisualStyleBackColor = true;
            this.radioUseTemplate.CheckedChanged += new System.EventHandler(this.UpdateButtons);
            // 
            // buttonBrowseTemplate
            // 
            this.buttonBrowseTemplate.Location = new System.Drawing.Point(369, 72);
            this.buttonBrowseTemplate.Name = "buttonBrowseTemplate";
            this.buttonBrowseTemplate.Size = new System.Drawing.Size(75, 23);
            this.buttonBrowseTemplate.TabIndex = 6;
            this.buttonBrowseTemplate.Text = "Browse...";
            this.buttonBrowseTemplate.UseVisualStyleBackColor = true;
            this.buttonBrowseTemplate.Click += new System.EventHandler(this.buttonBrowseTemplate_Click);
            // 
            // textBoxTemplatePath
            // 
            this.textBoxTemplatePath.Location = new System.Drawing.Point(30, 74);
            this.textBoxTemplatePath.Name = "textBoxTemplatePath";
            this.textBoxTemplatePath.ReadOnly = true;
            this.textBoxTemplatePath.Size = new System.Drawing.Size(333, 20);
            this.textBoxTemplatePath.TabIndex = 5;
            // 
            // checkAddVisioTemplates
            // 
            this.checkAddVisioTemplates.AutoSize = true;
            this.checkAddVisioTemplates.Location = new System.Drawing.Point(6, 5);
            this.checkAddVisioTemplates.Name = "checkAddVisioTemplates";
            this.checkAddVisioTemplates.Size = new System.Drawing.Size(177, 17);
            this.checkAddVisioTemplates.TabIndex = 2;
            this.checkAddVisioTemplates.Text = "Include custom Visio template(s)";
            this.checkAddVisioTemplates.UseVisualStyleBackColor = true;
            this.checkAddVisioTemplates.CheckedChanged += new System.EventHandler(this.UpdateButtons);
            this.checkAddVisioTemplates.Click += new System.EventHandler(this.UpdateButtons);
            // 
            // checkWixSetupDescription
            // 
            this.checkWixSetupDescription.ForeColor = System.Drawing.SystemColors.GrayText;
            this.checkWixSetupDescription.LinkArea = new System.Windows.Forms.LinkArea(196, 21);
            this.checkWixSetupDescription.LinkColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(192)))));
            this.checkWixSetupDescription.Location = new System.Drawing.Point(32, 104);
            this.checkWixSetupDescription.Name = "checkWixSetupDescription";
            this.checkWixSetupDescription.Size = new System.Drawing.Size(513, 40);
            this.checkWixSetupDescription.TabIndex = 1;
            this.checkWixSetupDescription.TabStop = true;
            this.checkWixSetupDescription.Text = resources.GetString("checkWixSetupDescription.Text");
            this.checkWixSetupDescription.UseCompatibleTextRendering = true;
            this.checkWixSetupDescription.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.checkWixSetupDescription_LinkClicked);
            // 
            // checkWixSetup
            // 
            this.checkWixSetup.AutoSize = true;
            this.checkWixSetup.Location = new System.Drawing.Point(16, 80);
            this.checkWixSetup.Name = "checkWixSetup";
            this.checkWixSetup.Size = new System.Drawing.Size(144, 17);
            this.checkWixSetup.TabIndex = 0;
            this.checkWixSetup.Text = "Create WiX setup project";
            this.checkWixSetup.UseVisualStyleBackColor = true;
            this.checkWixSetup.Click += new System.EventHandler(this.UpdateButtons);
            // 
            // WizardForm
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            this.ClientSize = new System.Drawing.Size(569, 462);
            this.Controls.Add(this.addinWizard);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "WizardForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Create New Visio Project";
            this.addinWizard.ResumeLayout(false);
            this.pageOptions.ResumeLayout(false);
            this.pageOptions.PerformLayout();
            this.pageCheck.ResumeLayout(false);
            this.pageCheck.PerformLayout();
            this.panelStencil.ResumeLayout(false);
            this.panelStencil.PerformLayout();
            this.panelTemplate.ResumeLayout(false);
            this.panelTemplate.PerformLayout();
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
            checkAddVisioStencils.Enabled = checkWixSetup.Checked;
            checkAddVisioTemplates.Enabled = checkWixSetup.Checked;

            radioCreateNewStencil.Enabled = checkAddVisioStencils.Checked;
            radioUseStencil.Enabled = checkAddVisioStencils.Checked;
            textBoxStencilPath.Enabled = checkAddVisioStencils.Checked && radioUseStencil.Checked;
            buttonBrowseStencil.Enabled = textBoxStencilPath.Enabled;

            radioCreateNewTemplate.Enabled = checkAddVisioTemplates.Checked;
            radioUseTemplate.Enabled = checkAddVisioTemplates.Checked;
            textBoxTemplatePath.Enabled = checkAddVisioTemplates.Checked && radioUseTemplate.Checked;
            buttonBrowseTemplate.Enabled = textBoxTemplatePath.Enabled;

            textBoxStencilPath.Text = StencilPaths == null ? ""
                : string.Join(" ", StencilPaths.Select(p => string.Format(@"""{0}""", System.IO.Path.GetFileName(p))));

            textBoxTemplatePath.Text = TemplatePaths == null ? ""
                : string.Join(" ", TemplatePaths.Select(p => string.Format(@"""{0}""", System.IO.Path.GetFileName(p))));

            addinWizard.NextEnabled = 
                !(checkWixSetup.Checked && checkAddVisioStencils.Checked && radioUseStencil.Checked && textBoxStencilPath.Text.Length == 0)
                && 
                !(checkWixSetup.Checked && checkAddVisioTemplates.Checked && radioUseTemplate.Checked && textBoxTemplatePath.Text.Length == 0)
                ;
        }
        
        private void checkWixSetupDescription_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Process.Start(new ProcessStartInfo("http://wixtoolset.org/"));
        }

        private void buttonBrowseTemplate_Click(object sender, EventArgs e)
        {
            var dlg = new OpenFileDialog
            {
                Filter = "Visio Template Files (*.vst, *.vtx, *.vst?)|*.vst;*.vtx;*.vst?", 
                Multiselect = true
            };

            if (dlg.ShowDialog() == DialogResult.OK)
            {
                TemplatePaths = dlg.FileNames;
                UpdateButtons(null, null);
            }
        }

        private void buttonBrowseStencil_Click(object sender, EventArgs e)
        {
            var dlg = new OpenFileDialog
            {
                Filter = "Visio Stencil Files (*.vst, *.vtx, *.vst?)|*.vss;*.vsx;*.vss?",
                Multiselect = true
            };

            if (dlg.ShowDialog() == DialogResult.OK)
            {
                StencilPaths = dlg.FileNames;
                UpdateButtons(null, null);
            }
        }
	}
}
