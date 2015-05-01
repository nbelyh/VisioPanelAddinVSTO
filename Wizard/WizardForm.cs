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
        public WizardForm()
        {
            // required for designer support
            InitializeComponent();

            checkSupportTaskPane.DataBindings.Add("Checked", this, "TaskPane");
            checkSupportRibbon.DataBindings.Add("Checked", this, "Ribbon");
            checkSupportCommandBars.DataBindings.Add("Checked", this, "CommandBars");
            checkWixSetup.DataBindings.Add("Checked", this, "GenerateWixSetup");
        }

        public bool TaskPane { get; set; }
        public bool Ribbon { get; set; }
        public bool CommandBars { get; set; }

        public bool GenerateWixSetup { get; set; }

	    public Image HeaderImage
	    {
	        get { return addinWizard.HeaderImage; }
            set { addinWizard.HeaderImage = value; }
	    }

		#region Designer generated code
        private WizardPage pageOptions;
        private Wizard addinWizard;
        private WizardPage pageCheck;
        private Label label5;
        private CheckBox checkSupportTaskPane;
        private Label label4;
        private Label label3;
        private CheckBox checkSupportCommandBars;
        private CheckBox checkSupportRibbon;
        private Label label2;
        private CheckBox checkWixSetup;
		
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
            this.pageOptions = new PanelAddinWizard.WizardPage();
            this.label5 = new System.Windows.Forms.Label();
            this.checkSupportTaskPane = new System.Windows.Forms.CheckBox();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.checkSupportCommandBars = new System.Windows.Forms.CheckBox();
            this.checkSupportRibbon = new System.Windows.Forms.CheckBox();
            this.addinWizard = new PanelAddinWizard.Wizard();
            this.pageCheck = new PanelAddinWizard.WizardPage();
            this.checkWixSetup = new System.Windows.Forms.CheckBox();
            this.label2 = new System.Windows.Forms.Label();
            this.pageOptions.SuspendLayout();
            this.addinWizard.SuspendLayout();
            this.pageCheck.SuspendLayout();
            this.SuspendLayout();
            // 
            // pageOptions
            // 
            this.pageOptions.Controls.Add(this.label5);
            this.pageOptions.Controls.Add(this.checkSupportTaskPane);
            this.pageOptions.Controls.Add(this.label4);
            this.pageOptions.Controls.Add(this.label3);
            this.pageOptions.Controls.Add(this.checkSupportCommandBars);
            this.pageOptions.Controls.Add(this.checkSupportRibbon);
            this.pageOptions.Description = "Please select add-in features";
            this.pageOptions.Location = new System.Drawing.Point(0, 0);
            this.pageOptions.Name = "pageOptions";
            this.pageOptions.Size = new System.Drawing.Size(466, 296);
            this.pageOptions.TabIndex = 11;
            this.pageOptions.Title = "Add-in options";
            // 
            // label5
            // 
            this.label5.ForeColor = System.Drawing.SystemColors.GrayText;
            this.label5.Location = new System.Drawing.Point(32, 104);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(380, 24);
            this.label5.TabIndex = 12;
            this.label5.Text = "Adds a docking panel which can be controlelled with a toggle button";
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
            // label4
            // 
            this.label4.ForeColor = System.Drawing.SystemColors.GrayText;
            this.label4.Location = new System.Drawing.Point(32, 200);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(380, 24);
            this.label4.TabIndex = 14;
            this.label4.Text = "Add a toolbar with custom images (check if you need to support old Visio)";
            // 
            // label3
            // 
            this.label3.ForeColor = System.Drawing.SystemColors.GrayText;
            this.label3.Location = new System.Drawing.Point(32, 152);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(380, 24);
            this.label3.TabIndex = 13;
            this.label3.Text = "Add a ribbon with buttons and custom images and state";
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
            this.pageCheck.Controls.Add(this.checkWixSetup);
            this.pageCheck.Controls.Add(this.label2);
            this.pageCheck.Description = "Please select deployment options";
            this.pageCheck.Location = new System.Drawing.Point(0, 0);
            this.pageCheck.Name = "pageCheck";
            this.pageCheck.Size = new System.Drawing.Size(466, 296);
            this.pageCheck.TabIndex = 13;
            this.pageCheck.Title = "Setup project";
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
            // 
            // label2
            // 
            this.label2.ForeColor = System.Drawing.SystemColors.GrayText;
            this.label2.Location = new System.Drawing.Point(32, 104);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(416, 32);
            this.label2.TabIndex = 18;
            this.label2.Text = "Create a WiX installer project to install addin for all users. If you skip the MS" +
    "I installer you can deploy your addin using ClickOnce publishing.";
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
            this.Text = "Sample Wizard";
            this.pageOptions.ResumeLayout(false);
            this.pageOptions.PerformLayout();
            this.addinWizard.ResumeLayout(false);
            this.pageCheck.ResumeLayout(false);
            this.pageCheck.PerformLayout();
            this.ResumeLayout(false);

        }
		#endregion

	    private void addinWizard_AfterSwitchPages(object sender, Wizard.AfterSwitchPagesEventArgs e)
		{
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
	}
}
