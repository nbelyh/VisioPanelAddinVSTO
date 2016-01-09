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
using System.IO;
using System.Linq;
using System.Windows.Forms;

#endregion

namespace PanelAddinWizard
{
	/// <summary>
	/// Summary description for WizardForm.
	/// </summary>
	public class WizardForm : Form
	{
	    private readonly IWizardFormHost _host;
        /// <summary>
        /// Creates a new instance of the <see cref="WizardForm"/> class.
        /// </summary>
        public WizardForm(IWizardFormHost host)
        {
            _host = host;

            // required for designer support
            InitializeComponent();

            checkWixSetup.Enabled = checkWixSetup.Checked = _host.IsWixInstalled();
            checkAddinProject.Enabled = checkAddinProject.Checked = _host.IsVstoInstalled();

            checkAddVisioFiles.Checked = false;
            radioCreateNewVisioFiles.Checked = true;
            checkCopyVisioFiles.Checked = true;

            checkEnableSetupUI.Checked = true;
            comboSetupUI.SelectedIndex = 4;

            UpdateButtons(null, null);
        }

        public bool TaskPane { get { return checkSupportTaskPane.Checked; } }
        public bool CommandBars { get { return checkSupportCommandBars.Checked; } }
        public bool AddinEnabled { get { return checkAddinProject.Checked; } }

        public bool RibbonXml { get { return checkSupportRibbon.Checked && radioSupportRibbonXml.Checked; } }
        public bool RibbonComponent { get { return checkSupportRibbon.Checked && radioSupportRibbonDesigner.Checked; } }

        public bool WixSetup { get { return checkWixSetup.Checked; } }

        public bool AddinTypeCOM { get { return radioAddinTypeCOM.Checked;  } }
        public bool AddinTypeVSTO { get { return radioAddinTypeVSTO.Checked; } }
        public string AddinName { get { return addinName.Text; } }
        public string AddinDescription { get { return addinDescription.Text; } }

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
        private ComboBox comboSetupUI;
        private CheckBox checkEnableSetupUI;
        private LinkLabel checkEnableSetupUIDescription;
        private OpenFileDialog licenseFileDialog;
        private WizardPage pageAddinOptions;
        private Label radioSupportRibbonXmlDescription;
        private Label radioSupportRibbonDesignerDescription;
        private RadioButton radioSupportRibbonXml;
        private RadioButton radioSupportRibbonDesigner;
        private Label checkSupportTaskPaneDescription;
        private CheckBox checkSupportTaskPane;
        private Label checkSupportCommandBarsDescription;
        private CheckBox checkSupportCommandBars;
        private CheckBox checkSupportRibbon;
        private Label addinDescriptionLabel;
        private Label addinNameLabel;
        private TextBox addinDescription;
        private TextBox addinName;
        private Label radioAddinTypeCOMLabel;
        private Label radioAddinTypeVSTOLabel;
        private RadioButton radioAddinTypeCOM;
        private RadioButton radioAddinTypeVSTO;
        private Label addinTypeLabel;

        private OpenFileDialog visioFileDialog;

        public WixSetupOptions GetSetupOptions()
        {
            return new WixSetupOptions
            {
                EnableWixSetup = checkWixSetup.Checked,
                AddVisioFiles = checkWixSetup.Checked && checkAddVisioFiles.Checked,
                CreateNewVisioFiles = checkAddVisioFiles.Checked && radioCreateNewVisioFiles.Checked,
                DuplicateExistingVisioFiles = checkAddVisioFiles.Checked && checkCopyVisioFiles.Checked,
                VisioFilePaths = visioFileDialog.FileNames,
                EnableWixUI = checkWixSetup.Checked && checkEnableSetupUI.Checked,
                WixUI = comboSetupUI.Text
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
        private LinkLabel checkWixSetupDescription;
        private CheckBox checkWixSetup;
		
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(WizardForm));
            this.visioFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.licenseFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.addinWizard = new PanelAddinWizard.Wizard();
            this.pageAddin = new PanelAddinWizard.WizardPage();
            this.addinTypeLabel = new System.Windows.Forms.Label();
            this.addinDescriptionLabel = new System.Windows.Forms.Label();
            this.addinNameLabel = new System.Windows.Forms.Label();
            this.addinDescription = new System.Windows.Forms.TextBox();
            this.addinName = new System.Windows.Forms.TextBox();
            this.radioAddinTypeCOMLabel = new System.Windows.Forms.Label();
            this.radioAddinTypeVSTOLabel = new System.Windows.Forms.Label();
            this.radioAddinTypeCOM = new System.Windows.Forms.RadioButton();
            this.radioAddinTypeVSTO = new System.Windows.Forms.RadioButton();
            this.checkAddinProjectDescription = new System.Windows.Forms.LinkLabel();
            this.checkAddinProject = new System.Windows.Forms.CheckBox();
            this.pageSetup = new PanelAddinWizard.WizardPage();
            this.checkEnableSetupUIDescription = new System.Windows.Forms.LinkLabel();
            this.comboSetupUI = new System.Windows.Forms.ComboBox();
            this.checkEnableSetupUI = new System.Windows.Forms.CheckBox();
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
            this.pageAddinOptions = new PanelAddinWizard.WizardPage();
            this.radioSupportRibbonXmlDescription = new System.Windows.Forms.Label();
            this.radioSupportRibbonDesignerDescription = new System.Windows.Forms.Label();
            this.radioSupportRibbonXml = new System.Windows.Forms.RadioButton();
            this.radioSupportRibbonDesigner = new System.Windows.Forms.RadioButton();
            this.checkSupportTaskPaneDescription = new System.Windows.Forms.Label();
            this.checkSupportTaskPane = new System.Windows.Forms.CheckBox();
            this.checkSupportCommandBarsDescription = new System.Windows.Forms.Label();
            this.checkSupportCommandBars = new System.Windows.Forms.CheckBox();
            this.checkSupportRibbon = new System.Windows.Forms.CheckBox();
            this.addinWizard.SuspendLayout();
            this.pageAddin.SuspendLayout();
            this.pageSetup.SuspendLayout();
            this.pageAddinOptions.SuspendLayout();
            this.SuspendLayout();
            // 
            // visioFileDialog
            // 
            this.visioFileDialog.Filter = "Visio templates and stencils|*.vst;*.vtx;*.vst?;*.vss;*.vsx;*.vss?|Visio template" +
    "s|*.vst;*.vtx;*.vst?|Visio stencils|*.vss;*.vsx;*.vss?|All Files|*.*";
            this.visioFileDialog.Multiselect = true;
            // 
            // licenseFileDialog
            // 
            this.licenseFileDialog.Filter = "Rich text format files|*.rtf|All Files|*.*";
            this.licenseFileDialog.Multiselect = true;
            // 
            // addinWizard
            // 
            this.addinWizard.Controls.Add(this.pageAddin);
            this.addinWizard.Controls.Add(this.pageSetup);
            this.addinWizard.Controls.Add(this.pageAddinOptions);
            this.addinWizard.HelpVisible = true;
            this.addinWizard.Location = new System.Drawing.Point(0, 0);
            this.addinWizard.Name = "addinWizard";
            this.addinWizard.Pages.AddRange(new PanelAddinWizard.WizardPage[] {
            this.pageAddin,
            this.pageAddinOptions,
            this.pageSetup});
            this.addinWizard.Size = new System.Drawing.Size(644, 514);
            this.addinWizard.TabIndex = 0;
            this.addinWizard.BeforeSwitchPages += new PanelAddinWizard.Wizard.BeforeSwitchPagesEventHandler(this.addinWizard_BeforeSwitchPages);
            this.addinWizard.AfterSwitchPages += new PanelAddinWizard.Wizard.AfterSwitchPagesEventHandler(this.addinWizard_AfterSwitchPages);
            this.addinWizard.Cancel += new System.ComponentModel.CancelEventHandler(this.addinWizard_Cancel);
            this.addinWizard.Finish += new System.ComponentModel.CancelEventHandler(this.addinWizard_Finish);
            this.addinWizard.Help += new System.EventHandler(this.addinWizard_Help);
            // 
            // pageAddin
            // 
            this.pageAddin.Controls.Add(this.addinTypeLabel);
            this.pageAddin.Controls.Add(this.addinDescriptionLabel);
            this.pageAddin.Controls.Add(this.addinNameLabel);
            this.pageAddin.Controls.Add(this.addinDescription);
            this.pageAddin.Controls.Add(this.addinName);
            this.pageAddin.Controls.Add(this.radioAddinTypeCOMLabel);
            this.pageAddin.Controls.Add(this.radioAddinTypeVSTOLabel);
            this.pageAddin.Controls.Add(this.radioAddinTypeCOM);
            this.pageAddin.Controls.Add(this.radioAddinTypeVSTO);
            this.pageAddin.Controls.Add(this.checkAddinProjectDescription);
            this.pageAddin.Controls.Add(this.checkAddinProject);
            this.pageAddin.Description = "Addin project general information";
            this.pageAddin.Location = new System.Drawing.Point(0, 0);
            this.pageAddin.Name = "pageAddin";
            this.pageAddin.Size = new System.Drawing.Size(644, 466);
            this.pageAddin.TabIndex = 11;
            this.pageAddin.Title = "Add-in project";
            // 
            // addinTypeLabel
            // 
            this.addinTypeLabel.AutoSize = true;
            this.addinTypeLabel.Location = new System.Drawing.Point(37, 303);
            this.addinTypeLabel.Name = "addinTypeLabel";
            this.addinTypeLabel.Size = new System.Drawing.Size(90, 13);
            this.addinTypeLabel.TabIndex = 38;
            this.addinTypeLabel.Text = "Type of the addin";
            // 
            // addinDescriptionLabel
            // 
            this.addinDescriptionLabel.AutoSize = true;
            this.addinDescriptionLabel.Location = new System.Drawing.Point(37, 218);
            this.addinDescriptionLabel.Name = "addinDescriptionLabel";
            this.addinDescriptionLabel.Size = new System.Drawing.Size(284, 13);
            this.addinDescriptionLabel.TabIndex = 37;
            this.addinDescriptionLabel.Text = "Addin description (leave blank to use AssemblyDescription)";
            // 
            // addinNameLabel
            // 
            this.addinNameLabel.AutoSize = true;
            this.addinNameLabel.Location = new System.Drawing.Point(37, 167);
            this.addinNameLabel.Name = "addinNameLabel";
            this.addinNameLabel.Size = new System.Drawing.Size(226, 13);
            this.addinNameLabel.TabIndex = 36;
            this.addinNameLabel.Text = "Addin name (leave blank to use AssemblyTitle)";
            // 
            // addinDescription
            // 
            this.addinDescription.Location = new System.Drawing.Point(40, 237);
            this.addinDescription.Multiline = true;
            this.addinDescription.Name = "addinDescription";
            this.addinDescription.Size = new System.Drawing.Size(436, 50);
            this.addinDescription.TabIndex = 35;
            // 
            // addinName
            // 
            this.addinName.Location = new System.Drawing.Point(40, 186);
            this.addinName.Name = "addinName";
            this.addinName.Size = new System.Drawing.Size(281, 20);
            this.addinName.TabIndex = 34;
            // 
            // radioAddinTypeCOMLabel
            // 
            this.radioAddinTypeCOMLabel.ForeColor = System.Drawing.SystemColors.GrayText;
            this.radioAddinTypeCOMLabel.Location = new System.Drawing.Point(75, 405);
            this.radioAddinTypeCOMLabel.Name = "radioAddinTypeCOMLabel";
            this.radioAddinTypeCOMLabel.Size = new System.Drawing.Size(401, 23);
            this.radioAddinTypeCOMLabel.TabIndex = 33;
            this.radioAddinTypeCOMLabel.Text = "Use this option to create COM addin. Does not use VSTO.";
            // 
            // radioAddinTypeVSTOLabel
            // 
            this.radioAddinTypeVSTOLabel.ForeColor = System.Drawing.SystemColors.GrayText;
            this.radioAddinTypeVSTOLabel.Location = new System.Drawing.Point(75, 351);
            this.radioAddinTypeVSTOLabel.Name = "radioAddinTypeVSTOLabel";
            this.radioAddinTypeVSTOLabel.Size = new System.Drawing.Size(401, 31);
            this.radioAddinTypeVSTOLabel.TabIndex = 31;
            this.radioAddinTypeVSTOLabel.Text = "Default. Creates the Addin project using Visual Studio Tools for Office. Select t" +
    "his option for visual studio designer support.";
            // 
            // radioAddinTypeCOM
            // 
            this.radioAddinTypeCOM.AutoSize = true;
            this.radioAddinTypeCOM.Location = new System.Drawing.Point(53, 385);
            this.radioAddinTypeCOM.Name = "radioAddinTypeCOM";
            this.radioAddinTypeCOM.Size = new System.Drawing.Size(157, 17);
            this.radioAddinTypeCOM.TabIndex = 32;
            this.radioAddinTypeCOM.Text = "Create a COM Addin project";
            this.radioAddinTypeCOM.UseVisualStyleBackColor = true;
            // 
            // radioAddinTypeVSTO
            // 
            this.radioAddinTypeVSTO.AutoSize = true;
            this.radioAddinTypeVSTO.Checked = true;
            this.radioAddinTypeVSTO.Location = new System.Drawing.Point(53, 331);
            this.radioAddinTypeVSTO.Name = "radioAddinTypeVSTO";
            this.radioAddinTypeVSTO.Size = new System.Drawing.Size(162, 17);
            this.radioAddinTypeVSTO.TabIndex = 30;
            this.radioAddinTypeVSTO.TabStop = true;
            this.radioAddinTypeVSTO.Text = "Create a VSTO Addin project";
            this.radioAddinTypeVSTO.UseVisualStyleBackColor = true;
            // 
            // checkAddinProjectDescription
            // 
            this.checkAddinProjectDescription.ForeColor = System.Drawing.SystemColors.GrayText;
            this.checkAddinProjectDescription.LinkArea = new System.Windows.Forms.LinkArea(229, 47);
            this.checkAddinProjectDescription.LinkColor = System.Drawing.SystemColors.HotTrack;
            this.checkAddinProjectDescription.Location = new System.Drawing.Point(40, 108);
            this.checkAddinProjectDescription.Name = "checkAddinProjectDescription";
            this.checkAddinProjectDescription.Size = new System.Drawing.Size(584, 46);
            this.checkAddinProjectDescription.TabIndex = 16;
            this.checkAddinProjectDescription.TabStop = true;
            this.checkAddinProjectDescription.Text = resources.GetString("checkAddinProjectDescription.Text");
            this.checkAddinProjectDescription.UseCompatibleTextRendering = true;
            this.checkAddinProjectDescription.VisitedLinkColor = System.Drawing.SystemColors.HotTrack;
            this.checkAddinProjectDescription.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.checkAddinProjectDescription_LinkClicked);
            // 
            // checkAddinProject
            // 
            this.checkAddinProject.AutoSize = true;
            this.checkAddinProject.Location = new System.Drawing.Point(24, 88);
            this.checkAddinProject.Name = "checkAddinProject";
            this.checkAddinProject.Size = new System.Drawing.Size(147, 17);
            this.checkAddinProject.TabIndex = 15;
            this.checkAddinProject.Text = "Create Visio Addin project";
            this.checkAddinProject.UseVisualStyleBackColor = true;
            this.checkAddinProject.CheckedChanged += new System.EventHandler(this.UpdateButtons);
            // 
            // pageSetup
            // 
            this.pageSetup.Controls.Add(this.checkEnableSetupUIDescription);
            this.pageSetup.Controls.Add(this.comboSetupUI);
            this.pageSetup.Controls.Add(this.checkEnableSetupUI);
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
            this.pageSetup.Size = new System.Drawing.Size(644, 466);
            this.pageSetup.TabIndex = 0;
            this.pageSetup.Title = "Setup project";
            // 
            // checkEnableSetupUIDescription
            // 
            this.checkEnableSetupUIDescription.ForeColor = System.Drawing.SystemColors.GrayText;
            this.checkEnableSetupUIDescription.LinkArea = new System.Windows.Forms.LinkArea(56, 13);
            this.checkEnableSetupUIDescription.LinkColor = System.Drawing.SystemColors.HotTrack;
            this.checkEnableSetupUIDescription.Location = new System.Drawing.Point(60, 350);
            this.checkEnableSetupUIDescription.Name = "checkEnableSetupUIDescription";
            this.checkEnableSetupUIDescription.Size = new System.Drawing.Size(516, 17);
            this.checkEnableSetupUIDescription.TabIndex = 15;
            this.checkEnableSetupUIDescription.TabStop = true;
            this.checkEnableSetupUIDescription.Text = "To learn more about user interface options, consult the documentation on the WiX " +
    "website.";
            this.checkEnableSetupUIDescription.UseCompatibleTextRendering = true;
            this.checkEnableSetupUIDescription.VisitedLinkColor = System.Drawing.SystemColors.HotTrack;
            this.checkEnableSetupUIDescription.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.checkEnableSetupUIDescription_LinkClicked);
            // 
            // comboSetupUI
            // 
            this.comboSetupUI.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboSetupUI.FormattingEnabled = true;
            this.comboSetupUI.Items.AddRange(new object[] {
            "WixUI_Minimal",
            "WixUI_InstallDir",
            "WixUI_InstallDirNoLicense",
            "WixUI_Mondo",
            "WixUI_Advanced"});
            this.comboSetupUI.Location = new System.Drawing.Point(57, 370);
            this.comboSetupUI.Name = "comboSetupUI";
            this.comboSetupUI.Size = new System.Drawing.Size(262, 21);
            this.comboSetupUI.TabIndex = 14;
            // 
            // checkEnableSetupUI
            // 
            this.checkEnableSetupUI.AutoSize = true;
            this.checkEnableSetupUI.Location = new System.Drawing.Point(41, 330);
            this.checkEnableSetupUI.Name = "checkEnableSetupUI";
            this.checkEnableSetupUI.Size = new System.Drawing.Size(200, 17);
            this.checkEnableSetupUI.TabIndex = 13;
            this.checkEnableSetupUI.Text = "Provide user interface for the installer";
            this.checkEnableSetupUI.UseVisualStyleBackColor = true;
            this.checkEnableSetupUI.Click += new System.EventHandler(this.UpdateButtons);
            // 
            // radioUseVisioFilesDescription
            // 
            this.radioUseVisioFilesDescription.ForeColor = System.Drawing.SystemColors.GrayText;
            this.radioUseVisioFilesDescription.Location = new System.Drawing.Point(82, 238);
            this.radioUseVisioFilesDescription.Name = "radioUseVisioFilesDescription";
            this.radioUseVisioFilesDescription.Size = new System.Drawing.Size(427, 15);
            this.radioUseVisioFilesDescription.TabIndex = 6;
            this.radioUseVisioFilesDescription.Text = "Choose this if you already have Visio templates(s) or stencil(s) to include it in" +
    " the project.";
            // 
            // checkAddVisioFiles
            // 
            this.checkAddVisioFiles.AutoSize = true;
            this.checkAddVisioFiles.Location = new System.Drawing.Point(41, 154);
            this.checkAddVisioFiles.Name = "checkAddVisioFiles";
            this.checkAddVisioFiles.Size = new System.Drawing.Size(206, 17);
            this.checkAddVisioFiles.TabIndex = 2;
            this.checkAddVisioFiles.Text = "Include Visio stencil or template files(s)";
            this.checkAddVisioFiles.UseVisualStyleBackColor = true;
            this.checkAddVisioFiles.CheckedChanged += new System.EventHandler(this.UpdateButtons);
            // 
            // radioCreateNewVisioFilesDescription
            // 
            this.radioCreateNewVisioFilesDescription.ForeColor = System.Drawing.SystemColors.GrayText;
            this.radioCreateNewVisioFilesDescription.Location = new System.Drawing.Point(85, 197);
            this.radioCreateNewVisioFilesDescription.Name = "radioCreateNewVisioFilesDescription";
            this.radioCreateNewVisioFilesDescription.Size = new System.Drawing.Size(540, 15);
            this.radioCreateNewVisioFilesDescription.TabIndex = 4;
            this.radioCreateNewVisioFilesDescription.Text = "Creates a sample template and stencil in the project. You could use them as a sta" +
    "rting point.";
            // 
            // checkWixSetupDescription
            // 
            this.checkWixSetupDescription.ForeColor = System.Drawing.SystemColors.GrayText;
            this.checkWixSetupDescription.LinkArea = new System.Windows.Forms.LinkArea(194, 21);
            this.checkWixSetupDescription.LinkColor = System.Drawing.SystemColors.HotTrack;
            this.checkWixSetupDescription.Location = new System.Drawing.Point(41, 106);
            this.checkWixSetupDescription.Name = "checkWixSetupDescription";
            this.checkWixSetupDescription.Size = new System.Drawing.Size(566, 34);
            this.checkWixSetupDescription.TabIndex = 1;
            this.checkWixSetupDescription.TabStop = true;
            this.checkWixSetupDescription.Text = resources.GetString("checkWixSetupDescription.Text");
            this.checkWixSetupDescription.UseCompatibleTextRendering = true;
            this.checkWixSetupDescription.VisitedLinkColor = System.Drawing.SystemColors.HotTrack;
            this.checkWixSetupDescription.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.checkWixSetupDescription_LinkClicked);
            // 
            // textBoxVisioFilesPath
            // 
            this.textBoxVisioFilesPath.Location = new System.Drawing.Point(85, 256);
            this.textBoxVisioFilesPath.Name = "textBoxVisioFilesPath";
            this.textBoxVisioFilesPath.ReadOnly = true;
            this.textBoxVisioFilesPath.Size = new System.Drawing.Size(395, 20);
            this.textBoxVisioFilesPath.TabIndex = 7;
            // 
            // checkWixSetup
            // 
            this.checkWixSetup.AutoSize = true;
            this.checkWixSetup.Location = new System.Drawing.Point(25, 86);
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
            this.checkCopyVisioFiles.Location = new System.Drawing.Point(85, 282);
            this.checkCopyVisioFiles.Name = "checkCopyVisioFiles";
            this.checkCopyVisioFiles.Size = new System.Drawing.Size(185, 17);
            this.checkCopyVisioFiles.TabIndex = 9;
            this.checkCopyVisioFiles.Text = "Copy file(s) to the project directory";
            this.checkCopyVisioFiles.UseVisualStyleBackColor = true;
            // 
            // radioUseVisioFiles
            // 
            this.radioUseVisioFiles.AutoSize = true;
            this.radioUseVisioFiles.Location = new System.Drawing.Point(60, 218);
            this.radioUseVisioFiles.Name = "radioUseVisioFiles";
            this.radioUseVisioFiles.Size = new System.Drawing.Size(224, 17);
            this.radioUseVisioFiles.TabIndex = 5;
            this.radioUseVisioFiles.TabStop = true;
            this.radioUseVisioFiles.Text = "Use already existing template/stencil file(s)";
            this.radioUseVisioFiles.UseVisualStyleBackColor = true;
            // 
            // buttonBrowseVisioFiles
            // 
            this.buttonBrowseVisioFiles.Location = new System.Drawing.Point(486, 259);
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
            this.radioCreateNewVisioFiles.Location = new System.Drawing.Point(60, 177);
            this.radioCreateNewVisioFiles.Name = "radioCreateNewVisioFiles";
            this.radioCreateNewVisioFiles.Size = new System.Drawing.Size(243, 17);
            this.radioCreateNewVisioFiles.TabIndex = 3;
            this.radioCreateNewVisioFiles.TabStop = true;
            this.radioCreateNewVisioFiles.Text = "Create a new (sample) template and stencil file";
            this.radioCreateNewVisioFiles.UseVisualStyleBackColor = true;
            this.radioCreateNewVisioFiles.CheckedChanged += new System.EventHandler(this.UpdateButtons);
            // 
            // pageAddinOptions
            // 
            this.pageAddinOptions.Controls.Add(this.radioSupportRibbonXmlDescription);
            this.pageAddinOptions.Controls.Add(this.radioSupportRibbonDesignerDescription);
            this.pageAddinOptions.Controls.Add(this.radioSupportRibbonXml);
            this.pageAddinOptions.Controls.Add(this.radioSupportRibbonDesigner);
            this.pageAddinOptions.Controls.Add(this.checkSupportTaskPaneDescription);
            this.pageAddinOptions.Controls.Add(this.checkSupportTaskPane);
            this.pageAddinOptions.Controls.Add(this.checkSupportCommandBarsDescription);
            this.pageAddinOptions.Controls.Add(this.checkSupportCommandBars);
            this.pageAddinOptions.Controls.Add(this.checkSupportRibbon);
            this.pageAddinOptions.Description = "Please select the add-in project features";
            this.pageAddinOptions.Location = new System.Drawing.Point(0, 0);
            this.pageAddinOptions.Name = "pageAddinOptions";
            this.pageAddinOptions.Size = new System.Drawing.Size(428, 208);
            this.pageAddinOptions.TabIndex = 12;
            this.pageAddinOptions.Title = "Add-in project options";
            // 
            // radioSupportRibbonXmlDescription
            // 
            this.radioSupportRibbonXmlDescription.ForeColor = System.Drawing.SystemColors.GrayText;
            this.radioSupportRibbonXmlDescription.Location = new System.Drawing.Point(67, 179);
            this.radioSupportRibbonXmlDescription.Name = "radioSupportRibbonXmlDescription";
            this.radioSupportRibbonXmlDescription.Size = new System.Drawing.Size(525, 51);
            this.radioSupportRibbonXmlDescription.TabIndex = 29;
            this.radioSupportRibbonXmlDescription.Text = resources.GetString("radioSupportRibbonXmlDescription.Text");
            // 
            // radioSupportRibbonDesignerDescription
            // 
            this.radioSupportRibbonDesignerDescription.ForeColor = System.Drawing.SystemColors.GrayText;
            this.radioSupportRibbonDesignerDescription.Location = new System.Drawing.Point(67, 134);
            this.radioSupportRibbonDesignerDescription.Name = "radioSupportRibbonDesignerDescription";
            this.radioSupportRibbonDesignerDescription.Size = new System.Drawing.Size(525, 22);
            this.radioSupportRibbonDesignerDescription.TabIndex = 27;
            this.radioSupportRibbonDesignerDescription.Text = "Adds the builtin visual ribbon designer component to the project";
            // 
            // radioSupportRibbonXml
            // 
            this.radioSupportRibbonXml.AutoSize = true;
            this.radioSupportRibbonXml.Checked = true;
            this.radioSupportRibbonXml.Location = new System.Drawing.Point(45, 159);
            this.radioSupportRibbonXml.Name = "radioSupportRibbonXml";
            this.radioSupportRibbonXml.Size = new System.Drawing.Size(159, 17);
            this.radioSupportRibbonXml.TabIndex = 28;
            this.radioSupportRibbonXml.TabStop = true;
            this.radioSupportRibbonXml.Text = "Design ribbon using XML file";
            this.radioSupportRibbonXml.UseVisualStyleBackColor = true;
            // 
            // radioSupportRibbonDesigner
            // 
            this.radioSupportRibbonDesigner.AutoSize = true;
            this.radioSupportRibbonDesigner.Location = new System.Drawing.Point(45, 114);
            this.radioSupportRibbonDesigner.Name = "radioSupportRibbonDesigner";
            this.radioSupportRibbonDesigner.Size = new System.Drawing.Size(135, 17);
            this.radioSupportRibbonDesigner.TabIndex = 26;
            this.radioSupportRibbonDesigner.Text = "Use the builtin designer";
            this.radioSupportRibbonDesigner.UseVisualStyleBackColor = true;
            // 
            // checkSupportTaskPaneDescription
            // 
            this.checkSupportTaskPaneDescription.ForeColor = System.Drawing.SystemColors.GrayText;
            this.checkSupportTaskPaneDescription.Location = new System.Drawing.Point(42, 264);
            this.checkSupportTaskPaneDescription.Name = "checkSupportTaskPaneDescription";
            this.checkSupportTaskPaneDescription.Size = new System.Drawing.Size(380, 24);
            this.checkSupportTaskPaneDescription.TabIndex = 24;
            this.checkSupportTaskPaneDescription.Text = "Adds a docking panel which can be controlelled with a toggle button";
            // 
            // checkSupportTaskPane
            // 
            this.checkSupportTaskPane.AutoSize = true;
            this.checkSupportTaskPane.Location = new System.Drawing.Point(26, 244);
            this.checkSupportTaskPane.Name = "checkSupportTaskPane";
            this.checkSupportTaskPane.Size = new System.Drawing.Size(157, 17);
            this.checkSupportTaskPane.TabIndex = 21;
            this.checkSupportTaskPane.Text = "Support Task Pane window";
            this.checkSupportTaskPane.UseVisualStyleBackColor = true;
            // 
            // checkSupportCommandBarsDescription
            // 
            this.checkSupportCommandBarsDescription.ForeColor = System.Drawing.SystemColors.GrayText;
            this.checkSupportCommandBarsDescription.Location = new System.Drawing.Point(42, 325);
            this.checkSupportCommandBarsDescription.Name = "checkSupportCommandBarsDescription";
            this.checkSupportCommandBarsDescription.Size = new System.Drawing.Size(380, 24);
            this.checkSupportCommandBarsDescription.TabIndex = 25;
            this.checkSupportCommandBarsDescription.Text = "Add a toolbar with custom images (check if you need to support old Visio)";
            // 
            // checkSupportCommandBars
            // 
            this.checkSupportCommandBars.AutoSize = true;
            this.checkSupportCommandBars.Location = new System.Drawing.Point(26, 305);
            this.checkSupportCommandBars.Name = "checkSupportCommandBars";
            this.checkSupportCommandBars.Size = new System.Drawing.Size(247, 17);
            this.checkSupportCommandBars.TabIndex = 23;
            this.checkSupportCommandBars.Text = "Support Command Bars (Visio 2007 and below)";
            this.checkSupportCommandBars.UseVisualStyleBackColor = true;
            // 
            // checkSupportRibbon
            // 
            this.checkSupportRibbon.AutoSize = true;
            this.checkSupportRibbon.Checked = true;
            this.checkSupportRibbon.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkSupportRibbon.Location = new System.Drawing.Point(26, 87);
            this.checkSupportRibbon.Name = "checkSupportRibbon";
            this.checkSupportRibbon.Size = new System.Drawing.Size(279, 17);
            this.checkSupportRibbon.TabIndex = 22;
            this.checkSupportRibbon.Text = "Support Ribbon user interface (Visio 2010 and above)";
            this.checkSupportRibbon.UseVisualStyleBackColor = true;
            // 
            // WizardForm
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            this.ClientSize = new System.Drawing.Size(644, 514);
            this.Controls.Add(this.addinWizard);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "WizardForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Create New Visio Project";
            this.addinWizard.ResumeLayout(false);
            this.pageAddin.ResumeLayout(false);
            this.pageAddin.PerformLayout();
            this.pageSetup.ResumeLayout(false);
            this.pageSetup.PerformLayout();
            this.pageAddinOptions.ResumeLayout(false);
            this.pageAddinOptions.PerformLayout();
            this.ResumeLayout(false);

        }
		#endregion

	    private void addinWizard_AfterSwitchPages(object sender, Wizard.AfterSwitchPagesEventArgs e)
		{
            UpdateButtons(null, null);
		}

		private void addinWizard_BeforeSwitchPages(object sender, Wizard.BeforeSwitchPagesEventArgs e)
		{
		    if (e.OldIndex == 0 && !checkAddinProject.Checked)
		        e.NewIndex = 2;

		    if (e.OldIndex == 2 && !checkAddinProject.Checked)
		        e.NewIndex = 0;
		}

		private void addinWizard_Cancel(object sender, CancelEventArgs e)
		{
        }

        private void addinWizard_Help(object sender, EventArgs e)
        {
        }

        private void UpdateButtons(object sender, EventArgs e)
        {
            addinName.Enabled = checkAddinProject.Checked;
            addinNameLabel.ForeColor = checkAddinProject.Checked ? SystemColors.ControlText : SystemColors.ControlDark;

            addinDescription.Enabled = checkAddinProject.Checked;
            addinDescriptionLabel.ForeColor = checkAddinProject.Checked ? SystemColors.ControlText : SystemColors.ControlDark;

            radioAddinTypeCOM.Enabled = checkAddinProject.Checked;
            radioAddinTypeCOMLabel.ForeColor = checkAddinProject.Checked ? SystemColors.GrayText : SystemColors.ControlDark;

            radioAddinTypeVSTO.Enabled = checkAddinProject.Checked;
            radioAddinTypeVSTOLabel.ForeColor = checkAddinProject.Checked ? SystemColors.GrayText : SystemColors.ControlDark;

            checkAddinProjectDescription.ForeColor = checkAddinProject.Checked ? SystemColors.GrayText : SystemColors.ControlDark;

            checkAddVisioFiles.Enabled = checkWixSetup.Checked;
            radioCreateNewVisioFiles.Enabled = checkWixSetup.Checked && checkAddVisioFiles.Checked;
            radioCreateNewVisioFilesDescription.ForeColor = radioCreateNewVisioFiles.Enabled ? SystemColors.GrayText : SystemColors.ControlDark;

            radioUseVisioFiles.Enabled = checkWixSetup.Checked && checkAddVisioFiles.Checked;
            radioUseVisioFilesDescription.ForeColor = radioUseVisioFiles.Enabled ? SystemColors.GrayText : SystemColors.ControlDark;

            textBoxVisioFilesPath.Enabled = checkWixSetup.Checked && checkAddVisioFiles.Checked && radioUseVisioFiles.Checked;

            buttonBrowseVisioFiles.Enabled = textBoxVisioFilesPath.Enabled;
            checkCopyVisioFiles.Enabled = textBoxVisioFilesPath.Enabled;

            textBoxVisioFilesPath.Text = visioFileDialog.FileNames == null ? ""
                : string.Join(" ", visioFileDialog.FileNames.Select(p => string.Format(@"""{0}""", Path.GetFileName(p))));

            checkSupportTaskPane.Enabled = checkAddinProject.Checked;
            checkSupportTaskPaneDescription.ForeColor = checkSupportTaskPane.Enabled ? SystemColors.GrayText : SystemColors.ControlDark;

            checkSupportRibbon.Enabled = checkAddinProject.Checked;

            radioSupportRibbonDesigner.Enabled = checkSupportRibbon.Enabled && checkSupportRibbon.Checked && radioAddinTypeVSTO.Checked;
            radioSupportRibbonDesignerDescription.ForeColor = radioSupportRibbonDesigner.Enabled ? SystemColors.GrayText : SystemColors.ControlDark;

            radioSupportRibbonXml.Enabled = checkSupportRibbon.Enabled && checkSupportRibbon.Checked;
            radioSupportRibbonXmlDescription.ForeColor = radioSupportRibbonXml.Enabled ? SystemColors.GrayText : SystemColors.ControlDark;

            checkSupportCommandBars.Enabled = checkAddinProject.Checked;
            checkSupportCommandBarsDescription.ForeColor = checkSupportCommandBars.Enabled ? SystemColors.GrayText : SystemColors.ControlDark;

            checkEnableSetupUI.Enabled = checkWixSetup.Checked;
            checkEnableSetupUIDescription.ForeColor = checkWixSetup.Checked ? SystemColors.GrayText : SystemColors.ControlDark;

            comboSetupUI.Enabled = checkWixSetup.Checked && checkEnableSetupUI.Checked;

            addinWizard.NextEnabled =
                !(checkWixSetup.Checked && checkAddVisioFiles.Checked && radioUseVisioFiles.Checked && textBoxVisioFilesPath.Text.Length == 0);
        }
        
        private void checkWixSetupDescription_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Process.Start(new ProcessStartInfo("http://wixtoolset.org/"));
        }

        private void buttonBrowse_Click(object sender, EventArgs e)
        {
            if (visioFileDialog.ShowDialog() == DialogResult.OK)
                UpdateButtons(null, null);
        }

        private void buttonLicenseFileBrowse_Click(object sender, EventArgs e)
        {
            if (licenseFileDialog.ShowDialog() == DialogResult.OK)
                UpdateButtons(null, null);
        }

        private void checkEnableSetupUIDescription_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            _host.OpenExternalLink(ExternalLink.WixDocsUI);
        }

        private void checkAddinProjectDescription_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            _host.OpenExternalLink(ExternalLink.VstoDownload);
        }

        private void addinWizard_Finish(object sender, CancelEventArgs e)
        {
            if (!checkAddinProject.Checked && !checkWixSetup.Checked)
            {
                MessageBox.Show(
                    "You have not selected any project(s) to create. \n\n" +
                    "Please select either the AddIn project, or Setup project, or combination of the two. " +
                    "If some projects creation options are disabled then please install the missing components (check download links links in descriptions).",
                    "Extended Visio Addin: Nothing to create", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                e.Cancel = true;
            }
        }
    }
}
