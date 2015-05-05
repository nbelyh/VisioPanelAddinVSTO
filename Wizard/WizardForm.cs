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

            checkSupportTaskPane.DataBindings.Add("Checked", this, "TaskPane");
            checkSupportRibbon.DataBindings.Add("Checked", this, "Ribbon");
            checkSupportCommandBars.DataBindings.Add("Checked", this, "CommandBars");

            checkWixSetup.DataBindings.Add("Checked", this, "WixSetup");

            checkAddVisioStencils.DataBindings.Add("Checked", this, "AddVisioStencils");
            checkAddVisioTemplates.DataBindings.Add("Checked", this, "AddVisioTemplates");

            checkWixSetup.Enabled = wixInstalled;

            UpdateButtons(null, null);
        }

        public bool TaskPane { get; set; }
        public bool Ribbon { get; set; }
        public bool CommandBars { get; set; }

        public bool WixSetup { get; set; }

        public bool AddVisioTemplates { get; set; }
        public string VisioTemplates { get; set; }

        public bool AddVisioStencils { get; set; }
	    public string VisioStencils { get; set; }

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
        private Label labelAddVisioStencils;
        private Label labelAddVisioTemplates;
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
            this.pageCheck = new PanelAddinWizard.WizardPage();
            this.labelAddVisioStencils = new System.Windows.Forms.Label();
            this.labelAddVisioTemplates = new System.Windows.Forms.Label();
            this.checkAddVisioStencils = new System.Windows.Forms.CheckBox();
            this.checkAddVisioTemplates = new System.Windows.Forms.CheckBox();
            this.checkWixSetup = new System.Windows.Forms.CheckBox();
            this.pageOptions = new PanelAddinWizard.WizardPage();
            this.labelSupportTaskPane = new System.Windows.Forms.Label();
            this.checkSupportTaskPane = new System.Windows.Forms.CheckBox();
            this.labelSupportCommandBars = new System.Windows.Forms.Label();
            this.labelSupportRibbon = new System.Windows.Forms.Label();
            this.checkSupportCommandBars = new System.Windows.Forms.CheckBox();
            this.checkSupportRibbon = new System.Windows.Forms.CheckBox();
            this.checkWixSetupDescription = new System.Windows.Forms.LinkLabel();
            this.addinWizard.SuspendLayout();
            this.pageCheck.SuspendLayout();
            this.pageOptions.SuspendLayout();
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
            this.addinWizard.Size = new System.Drawing.Size(466, 344);
            this.addinWizard.TabIndex = 0;
            this.addinWizard.BeforeSwitchPages += new PanelAddinWizard.Wizard.BeforeSwitchPagesEventHandler(this.addinWizard_BeforeSwitchPages);
            this.addinWizard.AfterSwitchPages += new PanelAddinWizard.Wizard.AfterSwitchPagesEventHandler(this.addinWizard_AfterSwitchPages);
            this.addinWizard.Cancel += new System.ComponentModel.CancelEventHandler(this.addinWizard_Cancel);
            this.addinWizard.Help += new System.EventHandler(this.addinWizard_Help);
            // 
            // pageCheck
            // 
            this.pageCheck.Controls.Add(this.labelAddVisioStencils);
            this.pageCheck.Controls.Add(this.labelAddVisioTemplates);
            this.pageCheck.Controls.Add(this.checkAddVisioStencils);
            this.pageCheck.Controls.Add(this.checkAddVisioTemplates);
            this.pageCheck.Controls.Add(this.checkWixSetupDescription);
            this.pageCheck.Controls.Add(this.checkWixSetup);
            this.pageCheck.Description = "Please select deployment options";
            this.pageCheck.Location = new System.Drawing.Point(0, 0);
            this.pageCheck.Name = "pageCheck";
            this.pageCheck.Size = new System.Drawing.Size(466, 296);
            this.pageCheck.TabIndex = 13;
            this.pageCheck.Title = "Setup project";
            // 
            // labelAddVisioStencils
            // 
            this.labelAddVisioStencils.ForeColor = System.Drawing.SystemColors.GrayText;
            this.labelAddVisioStencils.Location = new System.Drawing.Point(56, 232);
            this.labelAddVisioStencils.Name = "labelAddVisioStencils";
            this.labelAddVisioStencils.Size = new System.Drawing.Size(380, 16);
            this.labelAddVisioStencils.TabIndex = 24;
            this.labelAddVisioStencils.Text = "Creates a custom Visio stencil file and inclues it in Setup.";
            // 
            // labelAddVisioTemplates
            // 
            this.labelAddVisioTemplates.ForeColor = System.Drawing.SystemColors.GrayText;
            this.labelAddVisioTemplates.Location = new System.Drawing.Point(56, 184);
            this.labelAddVisioTemplates.Name = "labelAddVisioTemplates";
            this.labelAddVisioTemplates.Size = new System.Drawing.Size(380, 16);
            this.labelAddVisioTemplates.TabIndex = 23;
            this.labelAddVisioTemplates.Text = "Creates a custom Visio template file and includes it in Setup.";
            // 
            // checkAddVisioStencils
            // 
            this.checkAddVisioStencils.AutoSize = true;
            this.checkAddVisioStencils.Location = new System.Drawing.Point(32, 208);
            this.checkAddVisioStencils.Name = "checkAddVisioStencils";
            this.checkAddVisioStencils.Size = new System.Drawing.Size(164, 17);
            this.checkAddVisioStencils.TabIndex = 22;
            this.checkAddVisioStencils.Text = "Include cusom Visio stencil(s)";
            this.checkAddVisioStencils.UseVisualStyleBackColor = true;
            this.checkAddVisioStencils.Click += new System.EventHandler(this.UpdateButtons);
            // 
            // checkAddVisioTemplates
            // 
            this.checkAddVisioTemplates.AutoSize = true;
            this.checkAddVisioTemplates.Location = new System.Drawing.Point(32, 160);
            this.checkAddVisioTemplates.Name = "checkAddVisioTemplates";
            this.checkAddVisioTemplates.Size = new System.Drawing.Size(177, 17);
            this.checkAddVisioTemplates.TabIndex = 21;
            this.checkAddVisioTemplates.Text = "Include custom Visio template(s)";
            this.checkAddVisioTemplates.UseVisualStyleBackColor = true;
            this.checkAddVisioTemplates.Click += new System.EventHandler(this.UpdateButtons);
            // 
            // checkWixSetup
            // 
            this.checkWixSetup.AutoSize = true;
            this.checkWixSetup.Location = new System.Drawing.Point(16, 80);
            this.checkWixSetup.Name = "checkWixSetup";
            this.checkWixSetup.Size = new System.Drawing.Size(144, 17);
            this.checkWixSetup.TabIndex = 19;
            this.checkWixSetup.Text = "Create WiX setup project";
            this.checkWixSetup.UseVisualStyleBackColor = true;
            this.checkWixSetup.Click += new System.EventHandler(this.UpdateButtons);
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
            this.pageOptions.Size = new System.Drawing.Size(466, 296);
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
            // checkWixSetupDescription
            // 
            this.checkWixSetupDescription.ForeColor = System.Drawing.SystemColors.GrayText;
            this.checkWixSetupDescription.LinkArea = new System.Windows.Forms.LinkArea(196, 21);
            this.checkWixSetupDescription.LinkColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(192)))));
            this.checkWixSetupDescription.Location = new System.Drawing.Point(32, 104);
            this.checkWixSetupDescription.Name = "checkWixSetupDescription";
            this.checkWixSetupDescription.Size = new System.Drawing.Size(424, 40);
            this.checkWixSetupDescription.TabIndex = 20;
            this.checkWixSetupDescription.TabStop = true;
            this.checkWixSetupDescription.Text = resources.GetString("checkWixSetupDescription.Text");
            this.checkWixSetupDescription.UseCompatibleTextRendering = true;
            this.checkWixSetupDescription.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.checkWixSetupDescription_LinkClicked);
            // 
            // WizardForm
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            this.ClientSize = new System.Drawing.Size(466, 344);
            this.Controls.Add(this.addinWizard);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "WizardForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Create New Visio Project";
            this.addinWizard.ResumeLayout(false);
            this.pageCheck.ResumeLayout(false);
            this.pageCheck.PerformLayout();
            this.pageOptions.ResumeLayout(false);
            this.pageOptions.PerformLayout();
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
        }
        
        private void checkWixSetupDescription_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Process.Start(new ProcessStartInfo("http://wixtoolset.org/"));
        }
	}
}
