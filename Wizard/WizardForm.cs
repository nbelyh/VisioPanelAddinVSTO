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
using PanelAddinWizard.Properties;
using Microsoft.Win32;

#endregion

namespace PanelAddinWizard
{
    /// <summary>
	/// Summary description for WizardForm.
	/// </summary>
	public class WizardForm : Form, IWizardOptions
    {
	    private readonly IWizardFormHost _host;

        /// <summary>
        /// Creates a new instance of the <see cref="WizardForm"/> class.
        /// </summary>
        public WizardForm(IWizardFormHost host, string projectName)
        {
            _host = host;

            // required for designer support
            InitializeComponent();

            InstallExtensibilityInterop = GacUtil.GetInstalledVersion("Extensibility") == 0;

            checkWixSetup.Checked = true;

            var isVstoInstalled = _host.IsVstoInstalled();
            if (!isVstoInstalled)
            {
                radioAddinTypeVSTOLabel.Text = "Creates a VSTO addin. To enable this option, please install VSTO tools. For VS2012 and above, you can download them for free from the Microsoft Visual Studio Office Developer web site";
                radioAddinTypeVSTOLabel.LinkArea = new LinkArea(134, 40);
                radioAddinTypeVSTOLabel.ForeColor = Color.RosyBrown;

                radioAddinTypeCOM.Checked = true;
                radioAddinTypeVSTO.Checked = false;
            }

            var isWixInstalled = _host.IsWixInstalled();
            checkWixSetup.Enabled = isWixInstalled;
            if (!isWixInstalled)
            {
                checkWixSetup.Enabled = false;
                checkWixSetup.Checked = false;
                checkWixSetupDescription.ForeColor = Color.RosyBrown;
                checkWixSetupDescription.LinkColor = Color.Sienna;
                if (isVstoInstalled)
                    checkWixSetupDescription.Text = "To create a WiX installer you need to download and install free WiX toolset from http://wixtoolset.org. If you skip the MSI installer you can still deploy your addin using ClickOnce publishing, but the setup features listed below will not be available.";
                else
                    checkWixSetupDescription.Text = "To create a WiX installer you need to download and install free WiX toolset from http://wixtoolset.org.";
                checkWixSetupDescription.LinkArea = new LinkArea(81, 21);
            }

            checkAddinProject.Checked = true;

            var visioInteropVersion = GetVisioInteropVersion();

            InstallVisioInterops = visioInteropVersion < 15;

            checkLocalReferencesLabel.Text =
                string.Format(Resources.WizardForm_Adds_local_copies_of_the_interop_assemblies, GetVisioVersionName(visioInteropVersion));

            checkLocalReferences.Checked = InstallVisioInterops;
            if (visioInteropVersion == 0)
            {
                checkLocalReferencesLabel.ForeColor = Color.RosyBrown;
                checkLocalReferencesLabel.Text = "Local interops must be installed, no relevant assemblies found in GAC. Can be the case if Visio or Office is not installed at this machine.";
                checkLocalReferences.Checked = true;
            }

            radioSupportRibbonDesigner.Checked = radioAddinTypeVSTO.Checked;

            checkAddVisioFiles.Checked = false;
            radioCreateNewVisioFiles.Checked = true;
            checkCopyVisioFiles.Checked = true;

            checkEnableSetupUI.Checked = true;
            comboSetupUI.SelectedIndex = 4;

            checkEnableSetupLanguage.Checked = false;
            comboSetupLanguage.SelectedIndex = comboSetupLanguage.Items.IndexOf("en-US");

            var currentCulture = System.Threading.Thread.CurrentThread.CurrentCulture.Name;
            var index = comboSetupLanguage.Items.IndexOf(currentCulture);
            if (index >= 0 && index != comboSetupLanguage.SelectedIndex)
            {
                comboSetupLanguage.SelectedIndex = index;
                checkEnableSetupLanguage.Checked = true;
            }

            addinName.Text = projectName;
            addinManufacturer.Text = GetDefaultManufacturerName();

            UpdateButtons(null, null);
        }

        string GetDefaultManufacturerName()
        {
            var osHive = Environment.Is64BitOperatingSystem 
                ? RegistryView.Registry64 
                : RegistryView.Default;

            using (var baseKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, osHive))
            {
                if (baseKey != null)
                {
                    using (var key = baseKey.OpenSubKey(@"Software\Microsoft\Windows NT\CurrentVersion"))
                    {
                        if (key != null)
                        {
                            var registeredOrganization = key.GetValue("RegisteredOrganization", "").ToString();
                            if (!string.IsNullOrEmpty(registeredOrganization))
                                return registeredOrganization;
                        }
                    }
                }
            }

            return Environment.UserName;
        }

        public enum ExternalLink
        {
            WixDownload,
            WixDocsUI,
            VstoDownload,
            WixDocLanguage,
        }

        public void OpenExternalLink(ExternalLink link)
        {
            var url = GetExternalLinkUrl(link);

            if (url != null)
                Process.Start(new ProcessStartInfo(url));
        }

        string GetExternalLinkUrl(ExternalLink link)
        {
            switch (link)
            {
                case ExternalLink.WixDocsUI:
                    return "http://wixtoolset.org/documentation/manual/v3/wixui/wixui_dialog_library.html";

                case ExternalLink.WixDocLanguage:
                    return "http://wixtoolset.org/documentation/manual/v3/howtos/ui_and_localization";

                case ExternalLink.WixDownload:
                    return "http://wixtoolset.org/releases/";

                case ExternalLink.VstoDownload:
                    switch (_host.GetVisualStudioVersion())
                    {
                        case 11:
                            return "http://aka.ms/OfficeDevToolsForVS2012";
                        case 12:
                            return "http://aka.ms/OfficeDevToolsForVS2013";
                        case 14:
                            return "http://aka.ms/OfficeDevToolsForVS2015";
                        default:
                            return "http://dev.office.com/";
                    }
            }

            return null;
        }

	    private static int GetVisioInteropVersion()
	    {
            var versions = new[]
	        {
	            GacUtil.GetInstalledVersion("Office"),
	            GacUtil.GetInstalledVersion("Microsoft.Office.Interop.Visio")
	        };

	        return versions.Min();
	    }

        static string GetVisioVersionName(int version)
        {
            switch (version)
            {
                case 11:
                    return "(Visio 2003)";
                case 12:
                    return "(Visio 2007)";
                case 14:
                    return "(Visio 2010)";
                case 15:
                    return "(Visio 2013)";
                case 16:
                    return "(Visio 2016)";
                default:
                    return "(none)";
            }
        }

        public bool InstallExtensibilityInterop { get; set; }
        public bool InstallVisioInterops { get; set; }

        public bool TaskPane { get { return checkSupportTaskPane.Checked; } }
        public bool CommandBars { get { return checkSupportCommandBars.Checked; } }
        public bool AddinEnabled { get { return checkAddinProject.Checked; } }

        public bool RibbonXml { get { return checkSupportRibbon.Checked && radioSupportRibbonXml.Checked; } }
        public bool RibbonComponent { get { return checkSupportRibbon.Checked && radioSupportRibbonDesigner.Checked; } }

        public bool AddinTypeCOM { get { return checkAddinProject.Checked && radioAddinTypeCOM.Checked;  } }
        public bool AddinTypeVSTO { get { return checkAddinProject.Checked && radioAddinTypeVSTO.Checked; } }

        public string AddinName { get { return addinName.Text; } }
        public string AddinManufacturer { get { return addinManufacturer.Text; } }
        public string AddinDescription { get { return addinDescription.Text; } }

        public bool UseSetupLanguage { get { return checkEnableSetupLanguage.Checked; } }
        public string SetupLanguage { get { return comboSetupLanguage.Text; } }

	    public bool EnableWixSetup { get { return checkWixSetup.Checked; } }

        public bool AddVisioFiles { get { return checkWixSetup.Checked && checkAddVisioFiles.Checked; } }
        public bool CreateNewVisioFiles { get { return checkAddVisioFiles.Checked && radioCreateNewVisioFiles.Checked; } }
        public bool DuplicateExistingVisioFiles { get {return checkAddVisioFiles.Checked && checkCopyVisioFiles.Checked; } }
        public string[] VisioFilePaths { get { return visioFileDialog.FileNames; } }

        public bool EnableWixUI { get { return checkWixSetup.Checked && checkEnableSetupUI.Checked; } }
        public string WixUI { get { return comboSetupUI.Text; } }

        private Label radioUseVisioFilesDescription;
        private CheckBox checkAddVisioFiles;
        private Label radioCreateNewVisioFilesDescription;
        private TextBox textBoxVisioFilesPath;
        private CheckBox checkCopyVisioFiles;
        private Button buttonBrowseVisioFiles;
        private RadioButton radioCreateNewVisioFiles;
        private RadioButton radioUseVisioFiles;
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
        private RadioButton radioAddinTypeCOM;
        private RadioButton radioAddinTypeVSTO;
        private Label addinTypeLabel;
        private Label checkLocalReferencesLabel;
        private CheckBox checkLocalReferences;
        private LinkLabel radioAddinTypeVSTOLabel;
        private Label checkAddinProjectDescription;

        private OpenFileDialog visioFileDialog;

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
        private LinkLabel checkEnableSetupLanguageDescription;
        private ComboBox comboSetupLanguage;
        private CheckBox checkEnableSetupLanguage;
        private Label addinManufacturerLabel;
        private TextBox addinManufacturer;
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
            this.pageSetup = new PanelAddinWizard.WizardPage();
            this.checkEnableSetupLanguageDescription = new System.Windows.Forms.LinkLabel();
            this.comboSetupLanguage = new System.Windows.Forms.ComboBox();
            this.checkEnableSetupLanguage = new System.Windows.Forms.CheckBox();
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
            this.checkLocalReferencesLabel = new System.Windows.Forms.Label();
            this.checkLocalReferences = new System.Windows.Forms.CheckBox();
            this.radioSupportRibbonXmlDescription = new System.Windows.Forms.Label();
            this.radioSupportRibbonDesignerDescription = new System.Windows.Forms.Label();
            this.radioSupportRibbonXml = new System.Windows.Forms.RadioButton();
            this.radioSupportRibbonDesigner = new System.Windows.Forms.RadioButton();
            this.checkSupportTaskPaneDescription = new System.Windows.Forms.Label();
            this.checkSupportTaskPane = new System.Windows.Forms.CheckBox();
            this.checkSupportCommandBarsDescription = new System.Windows.Forms.Label();
            this.checkSupportCommandBars = new System.Windows.Forms.CheckBox();
            this.checkSupportRibbon = new System.Windows.Forms.CheckBox();
            this.pageAddin = new PanelAddinWizard.WizardPage();
            this.checkAddinProjectDescription = new System.Windows.Forms.Label();
            this.radioAddinTypeVSTOLabel = new System.Windows.Forms.LinkLabel();
            this.addinTypeLabel = new System.Windows.Forms.Label();
            this.addinDescriptionLabel = new System.Windows.Forms.Label();
            this.addinManufacturerLabel = new System.Windows.Forms.Label();
            this.addinNameLabel = new System.Windows.Forms.Label();
            this.addinManufacturer = new System.Windows.Forms.TextBox();
            this.addinDescription = new System.Windows.Forms.TextBox();
            this.addinName = new System.Windows.Forms.TextBox();
            this.radioAddinTypeCOMLabel = new System.Windows.Forms.Label();
            this.radioAddinTypeCOM = new System.Windows.Forms.RadioButton();
            this.radioAddinTypeVSTO = new System.Windows.Forms.RadioButton();
            this.checkAddinProject = new System.Windows.Forms.CheckBox();
            this.addinWizard.SuspendLayout();
            this.pageSetup.SuspendLayout();
            this.pageAddinOptions.SuspendLayout();
            this.pageAddin.SuspendLayout();
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
            this.addinWizard.Controls.Add(this.pageSetup);
            this.addinWizard.Controls.Add(this.pageAddinOptions);
            this.addinWizard.Controls.Add(this.pageAddin);
            this.addinWizard.HelpVisible = true;
            this.addinWizard.Location = new System.Drawing.Point(0, 0);
            this.addinWizard.Name = "addinWizard";
            this.addinWizard.Pages.AddRange(new PanelAddinWizard.WizardPage[] {
            this.pageAddin,
            this.pageAddinOptions,
            this.pageSetup});
            this.addinWizard.Size = new System.Drawing.Size(697, 554);
            this.addinWizard.TabIndex = 0;
            this.addinWizard.BeforeSwitchPages += new PanelAddinWizard.Wizard.BeforeSwitchPagesEventHandler(this.addinWizard_BeforeSwitchPages);
            this.addinWizard.AfterSwitchPages += new PanelAddinWizard.Wizard.AfterSwitchPagesEventHandler(this.addinWizard_AfterSwitchPages);
            this.addinWizard.Cancel += new System.ComponentModel.CancelEventHandler(this.addinWizard_Cancel);
            this.addinWizard.Finish += new System.ComponentModel.CancelEventHandler(this.addinWizard_Finish);
            this.addinWizard.Help += new System.EventHandler(this.addinWizard_Help);
            // 
            // pageSetup
            // 
            this.pageSetup.Controls.Add(this.checkEnableSetupLanguageDescription);
            this.pageSetup.Controls.Add(this.comboSetupLanguage);
            this.pageSetup.Controls.Add(this.checkEnableSetupLanguage);
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
            this.pageSetup.Size = new System.Drawing.Size(697, 506);
            this.pageSetup.TabIndex = 0;
            this.pageSetup.Title = "Setup project";
            // 
            // checkEnableSetupLanguageDescription
            // 
            this.checkEnableSetupLanguageDescription.ForeColor = System.Drawing.SystemColors.GrayText;
            this.checkEnableSetupLanguageDescription.LinkArea = new System.Windows.Forms.LinkArea(137, 13);
            this.checkEnableSetupLanguageDescription.LinkColor = System.Drawing.SystemColors.HotTrack;
            this.checkEnableSetupLanguageDescription.Location = new System.Drawing.Point(59, 449);
            this.checkEnableSetupLanguageDescription.Name = "checkEnableSetupLanguageDescription";
            this.checkEnableSetupLanguageDescription.Size = new System.Drawing.Size(516, 31);
            this.checkEnableSetupLanguageDescription.TabIndex = 14;
            this.checkEnableSetupLanguageDescription.TabStop = true;
            this.checkEnableSetupLanguageDescription.Text = "Please select the language this if you install files using non-english characters" +
    ". To learn more about localization options, consult the documentation on the WiX" +
    " website.";
            this.checkEnableSetupLanguageDescription.UseCompatibleTextRendering = true;
            this.checkEnableSetupLanguageDescription.VisitedLinkColor = System.Drawing.SystemColors.HotTrack;
            this.checkEnableSetupLanguageDescription.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.checkEnableSetupLanguageDescription_LinkClicked);
            // 
            // comboSetupLanguage
            // 
            this.comboSetupLanguage.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboSetupLanguage.FormattingEnabled = true;
            this.comboSetupLanguage.Items.AddRange(new object[] {
            "ar-SA",
            "bg-BG",
            "ca-ES",
            "cs-CZ",
            "da-DK",
            "de-DE",
            "el-GR",
            "en-US",
            "es-ES",
            "et-EE",
            "fi-FI",
            "fr-FR",
            "he-IL",
            "hi-IN",
            "hr-HR",
            "hu-HU",
            "it-IT",
            "ja-JP",
            "kk-KZ",
            "ko-KR",
            "lt-LT",
            "lv-LV",
            "nb-NO",
            "nl-NL",
            "pl-PL",
            "pt-BR",
            "pt-PT",
            "ro-RO",
            "ru-RU",
            "sk-SK",
            "sl-SI",
            "sr-LA",
            "sv-SE",
            "th-TH",
            "tr-TR",
            "uk-UA",
            "zh-CN",
            "zh-HK",
            "zh-TW"});
            this.comboSetupLanguage.Location = new System.Drawing.Point(59, 423);
            this.comboSetupLanguage.Name = "comboSetupLanguage";
            this.comboSetupLanguage.Size = new System.Drawing.Size(262, 21);
            this.comboSetupLanguage.TabIndex = 13;
            // 
            // checkEnableSetupLanguage
            // 
            this.checkEnableSetupLanguage.AutoSize = true;
            this.checkEnableSetupLanguage.Location = new System.Drawing.Point(40, 400);
            this.checkEnableSetupLanguage.Name = "checkEnableSetupLanguage";
            this.checkEnableSetupLanguage.Size = new System.Drawing.Size(334, 17);
            this.checkEnableSetupLanguage.TabIndex = 12;
            this.checkEnableSetupLanguage.Text = "Specify user interface language for the installer (default is English)";
            this.checkEnableSetupLanguage.UseVisualStyleBackColor = true;
            this.checkEnableSetupLanguage.CheckedChanged += new System.EventHandler(this.UpdateButtons);
            // 
            // checkEnableSetupUIDescription
            // 
            this.checkEnableSetupUIDescription.ForeColor = System.Drawing.SystemColors.GrayText;
            this.checkEnableSetupUIDescription.LinkArea = new System.Windows.Forms.LinkArea(56, 13);
            this.checkEnableSetupUIDescription.LinkColor = System.Drawing.SystemColors.HotTrack;
            this.checkEnableSetupUIDescription.Location = new System.Drawing.Point(59, 369);
            this.checkEnableSetupUIDescription.Name = "checkEnableSetupUIDescription";
            this.checkEnableSetupUIDescription.Size = new System.Drawing.Size(516, 16);
            this.checkEnableSetupUIDescription.TabIndex = 11;
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
            this.comboSetupUI.Location = new System.Drawing.Point(59, 345);
            this.comboSetupUI.Name = "comboSetupUI";
            this.comboSetupUI.Size = new System.Drawing.Size(262, 21);
            this.comboSetupUI.TabIndex = 10;
            // 
            // checkEnableSetupUI
            // 
            this.checkEnableSetupUI.AutoSize = true;
            this.checkEnableSetupUI.Location = new System.Drawing.Point(40, 322);
            this.checkEnableSetupUI.Name = "checkEnableSetupUI";
            this.checkEnableSetupUI.Size = new System.Drawing.Size(200, 17);
            this.checkEnableSetupUI.TabIndex = 9;
            this.checkEnableSetupUI.Text = "Provide user interface for the installer";
            this.checkEnableSetupUI.UseVisualStyleBackColor = true;
            this.checkEnableSetupUI.Click += new System.EventHandler(this.UpdateButtons);
            // 
            // radioUseVisioFilesDescription
            // 
            this.radioUseVisioFilesDescription.ForeColor = System.Drawing.SystemColors.GrayText;
            this.radioUseVisioFilesDescription.Location = new System.Drawing.Point(81, 238);
            this.radioUseVisioFilesDescription.Name = "radioUseVisioFilesDescription";
            this.radioUseVisioFilesDescription.Size = new System.Drawing.Size(427, 15);
            this.radioUseVisioFilesDescription.TabIndex = 6;
            this.radioUseVisioFilesDescription.Text = "Choose this if you already have Visio templates(s) or stencil(s) to include it in" +
    " the project.";
            // 
            // checkAddVisioFiles
            // 
            this.checkAddVisioFiles.AutoSize = true;
            this.checkAddVisioFiles.Location = new System.Drawing.Point(40, 154);
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
            this.radioCreateNewVisioFilesDescription.Location = new System.Drawing.Point(84, 197);
            this.radioCreateNewVisioFilesDescription.Name = "radioCreateNewVisioFilesDescription";
            this.radioCreateNewVisioFilesDescription.Size = new System.Drawing.Size(540, 15);
            this.radioCreateNewVisioFilesDescription.TabIndex = 4;
            this.radioCreateNewVisioFilesDescription.Text = "Creates a sample template and stencil in the project. You could use them as a sta" +
    "rting point.";
            // 
            // checkWixSetupDescription
            // 
            this.checkWixSetupDescription.ForeColor = System.Drawing.SystemColors.GrayText;
            this.checkWixSetupDescription.LinkArea = new System.Windows.Forms.LinkArea(170, 21);
            this.checkWixSetupDescription.LinkColor = System.Drawing.SystemColors.HotTrack;
            this.checkWixSetupDescription.Location = new System.Drawing.Point(40, 112);
            this.checkWixSetupDescription.Name = "checkWixSetupDescription";
            this.checkWixSetupDescription.Size = new System.Drawing.Size(566, 43);
            this.checkWixSetupDescription.TabIndex = 1;
            this.checkWixSetupDescription.TabStop = true;
            this.checkWixSetupDescription.Text = "Create a WiX installer project to install addin for all users. If you skip the MS" +
    "I installer you can deploy your addin using ClickOnce publishing. Lear more abou" +
    "t WIX at http://wixtoolset.org";
            this.checkWixSetupDescription.UseCompatibleTextRendering = true;
            this.checkWixSetupDescription.VisitedLinkColor = System.Drawing.SystemColors.HotTrack;
            this.checkWixSetupDescription.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.checkWixSetupDescription_LinkClicked);
            // 
            // textBoxVisioFilesPath
            // 
            this.textBoxVisioFilesPath.Location = new System.Drawing.Point(84, 258);
            this.textBoxVisioFilesPath.Name = "textBoxVisioFilesPath";
            this.textBoxVisioFilesPath.ReadOnly = true;
            this.textBoxVisioFilesPath.Size = new System.Drawing.Size(441, 20);
            this.textBoxVisioFilesPath.TabIndex = 7;
            // 
            // checkWixSetup
            // 
            this.checkWixSetup.AutoSize = true;
            this.checkWixSetup.Location = new System.Drawing.Point(24, 88);
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
            this.checkCopyVisioFiles.Location = new System.Drawing.Point(84, 284);
            this.checkCopyVisioFiles.Name = "checkCopyVisioFiles";
            this.checkCopyVisioFiles.Size = new System.Drawing.Size(185, 17);
            this.checkCopyVisioFiles.TabIndex = 8;
            this.checkCopyVisioFiles.Text = "Copy file(s) to the project directory";
            this.checkCopyVisioFiles.UseVisualStyleBackColor = true;
            // 
            // radioUseVisioFiles
            // 
            this.radioUseVisioFiles.AutoSize = true;
            this.radioUseVisioFiles.Location = new System.Drawing.Point(59, 218);
            this.radioUseVisioFiles.Name = "radioUseVisioFiles";
            this.radioUseVisioFiles.Size = new System.Drawing.Size(224, 17);
            this.radioUseVisioFiles.TabIndex = 5;
            this.radioUseVisioFiles.TabStop = true;
            this.radioUseVisioFiles.Text = "Use already existing template/stencil file(s)";
            this.radioUseVisioFiles.UseVisualStyleBackColor = true;
            // 
            // buttonBrowseVisioFiles
            // 
            this.buttonBrowseVisioFiles.Location = new System.Drawing.Point(531, 256);
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
            this.radioCreateNewVisioFiles.Location = new System.Drawing.Point(59, 177);
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
            this.pageAddinOptions.Controls.Add(this.checkLocalReferencesLabel);
            this.pageAddinOptions.Controls.Add(this.checkLocalReferences);
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
            this.pageAddinOptions.Size = new System.Drawing.Size(697, 506);
            this.pageAddinOptions.TabIndex = 0;
            this.pageAddinOptions.Title = "Add-in project options";
            // 
            // checkLocalReferencesLabel
            // 
            this.checkLocalReferencesLabel.ForeColor = System.Drawing.SystemColors.GrayText;
            this.checkLocalReferencesLabel.Location = new System.Drawing.Point(48, 408);
            this.checkLocalReferencesLabel.Name = "checkLocalReferencesLabel";
            this.checkLocalReferencesLabel.Size = new System.Drawing.Size(556, 36);
            this.checkLocalReferencesLabel.TabIndex = 10;
            this.checkLocalReferencesLabel.Text = "Adds local copies of the interop assemblies";
            // 
            // checkLocalReferences
            // 
            this.checkLocalReferences.AutoSize = true;
            this.checkLocalReferences.Location = new System.Drawing.Point(24, 384);
            this.checkLocalReferences.Name = "checkLocalReferences";
            this.checkLocalReferences.Size = new System.Drawing.Size(211, 17);
            this.checkLocalReferences.TabIndex = 9;
            this.checkLocalReferences.Text = "Add local Visio 2013 interop assemblies";
            this.checkLocalReferences.UseVisualStyleBackColor = true;
            // 
            // radioSupportRibbonXmlDescription
            // 
            this.radioSupportRibbonXmlDescription.ForeColor = System.Drawing.SystemColors.GrayText;
            this.radioSupportRibbonXmlDescription.Location = new System.Drawing.Point(64, 192);
            this.radioSupportRibbonXmlDescription.Name = "radioSupportRibbonXmlDescription";
            this.radioSupportRibbonXmlDescription.Size = new System.Drawing.Size(525, 48);
            this.radioSupportRibbonXmlDescription.TabIndex = 4;
            this.radioSupportRibbonXmlDescription.Text = resources.GetString("radioSupportRibbonXmlDescription.Text");
            // 
            // radioSupportRibbonDesignerDescription
            // 
            this.radioSupportRibbonDesignerDescription.ForeColor = System.Drawing.SystemColors.GrayText;
            this.radioSupportRibbonDesignerDescription.Location = new System.Drawing.Point(64, 136);
            this.radioSupportRibbonDesignerDescription.Name = "radioSupportRibbonDesignerDescription";
            this.radioSupportRibbonDesignerDescription.Size = new System.Drawing.Size(525, 32);
            this.radioSupportRibbonDesignerDescription.TabIndex = 2;
            this.radioSupportRibbonDesignerDescription.Text = "Adds the builtin visual ribbon designer component to the project. You will be abl" +
    "e to design your ribbon with the built-in Visual Studio ribbon designer.";
            // 
            // radioSupportRibbonXml
            // 
            this.radioSupportRibbonXml.AutoSize = true;
            this.radioSupportRibbonXml.Checked = true;
            this.radioSupportRibbonXml.Location = new System.Drawing.Point(40, 168);
            this.radioSupportRibbonXml.Name = "radioSupportRibbonXml";
            this.radioSupportRibbonXml.Size = new System.Drawing.Size(159, 17);
            this.radioSupportRibbonXml.TabIndex = 3;
            this.radioSupportRibbonXml.TabStop = true;
            this.radioSupportRibbonXml.Text = "Design ribbon using XML file";
            this.radioSupportRibbonXml.UseVisualStyleBackColor = true;
            // 
            // radioSupportRibbonDesigner
            // 
            this.radioSupportRibbonDesigner.AutoSize = true;
            this.radioSupportRibbonDesigner.Location = new System.Drawing.Point(40, 112);
            this.radioSupportRibbonDesigner.Name = "radioSupportRibbonDesigner";
            this.radioSupportRibbonDesigner.Size = new System.Drawing.Size(135, 17);
            this.radioSupportRibbonDesigner.TabIndex = 1;
            this.radioSupportRibbonDesigner.Text = "Use the builtin designer";
            this.radioSupportRibbonDesigner.UseVisualStyleBackColor = true;
            // 
            // checkSupportTaskPaneDescription
            // 
            this.checkSupportTaskPaneDescription.ForeColor = System.Drawing.SystemColors.GrayText;
            this.checkSupportTaskPaneDescription.Location = new System.Drawing.Point(48, 264);
            this.checkSupportTaskPaneDescription.Name = "checkSupportTaskPaneDescription";
            this.checkSupportTaskPaneDescription.Size = new System.Drawing.Size(536, 32);
            this.checkSupportTaskPaneDescription.TabIndex = 6;
            this.checkSupportTaskPaneDescription.Text = "Adds a sample docking panel which can be controlelled with a toggle button. You c" +
    "an customize this panel to show your controls.";
            // 
            // checkSupportTaskPane
            // 
            this.checkSupportTaskPane.AutoSize = true;
            this.checkSupportTaskPane.Location = new System.Drawing.Point(24, 248);
            this.checkSupportTaskPane.Name = "checkSupportTaskPane";
            this.checkSupportTaskPane.Size = new System.Drawing.Size(157, 17);
            this.checkSupportTaskPane.TabIndex = 5;
            this.checkSupportTaskPane.Text = "Support Task Pane window";
            this.checkSupportTaskPane.UseVisualStyleBackColor = true;
            // 
            // checkSupportCommandBarsDescription
            // 
            this.checkSupportCommandBarsDescription.ForeColor = System.Drawing.SystemColors.GrayText;
            this.checkSupportCommandBarsDescription.Location = new System.Drawing.Point(48, 336);
            this.checkSupportCommandBarsDescription.Name = "checkSupportCommandBarsDescription";
            this.checkSupportCommandBarsDescription.Size = new System.Drawing.Size(488, 32);
            this.checkSupportCommandBarsDescription.TabIndex = 8;
            this.checkSupportCommandBarsDescription.Text = "Add a toolbar with custom images. Only recommended if you need to support old ver" +
    "sion of the Visio (unchecked by default)";
            // 
            // checkSupportCommandBars
            // 
            this.checkSupportCommandBars.AutoSize = true;
            this.checkSupportCommandBars.Location = new System.Drawing.Point(24, 312);
            this.checkSupportCommandBars.Name = "checkSupportCommandBars";
            this.checkSupportCommandBars.Size = new System.Drawing.Size(247, 17);
            this.checkSupportCommandBars.TabIndex = 7;
            this.checkSupportCommandBars.Text = "Support Command Bars (Visio 2007 and below)";
            this.checkSupportCommandBars.UseVisualStyleBackColor = true;
            // 
            // checkSupportRibbon
            // 
            this.checkSupportRibbon.AutoSize = true;
            this.checkSupportRibbon.Checked = true;
            this.checkSupportRibbon.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkSupportRibbon.Location = new System.Drawing.Point(24, 88);
            this.checkSupportRibbon.Name = "checkSupportRibbon";
            this.checkSupportRibbon.Size = new System.Drawing.Size(279, 17);
            this.checkSupportRibbon.TabIndex = 0;
            this.checkSupportRibbon.Text = "Support Ribbon user interface (Visio 2010 and above)";
            this.checkSupportRibbon.UseVisualStyleBackColor = true;
            this.checkSupportRibbon.CheckedChanged += new System.EventHandler(this.UpdateButtons);
            // 
            // pageAddin
            // 
            this.pageAddin.Controls.Add(this.checkAddinProjectDescription);
            this.pageAddin.Controls.Add(this.radioAddinTypeVSTOLabel);
            this.pageAddin.Controls.Add(this.addinTypeLabel);
            this.pageAddin.Controls.Add(this.addinDescriptionLabel);
            this.pageAddin.Controls.Add(this.addinManufacturerLabel);
            this.pageAddin.Controls.Add(this.addinNameLabel);
            this.pageAddin.Controls.Add(this.addinManufacturer);
            this.pageAddin.Controls.Add(this.addinDescription);
            this.pageAddin.Controls.Add(this.addinName);
            this.pageAddin.Controls.Add(this.radioAddinTypeCOMLabel);
            this.pageAddin.Controls.Add(this.radioAddinTypeCOM);
            this.pageAddin.Controls.Add(this.radioAddinTypeVSTO);
            this.pageAddin.Controls.Add(this.checkAddinProject);
            this.pageAddin.Description = "Addin project general information";
            this.pageAddin.Location = new System.Drawing.Point(0, 0);
            this.pageAddin.Name = "pageAddin";
            this.pageAddin.Size = new System.Drawing.Size(697, 506);
            this.pageAddin.TabIndex = 0;
            this.pageAddin.Title = "Add-in project";
            // 
            // checkAddinProjectDescription
            // 
            this.checkAddinProjectDescription.ForeColor = System.Drawing.SystemColors.GrayText;
            this.checkAddinProjectDescription.Location = new System.Drawing.Point(37, 108);
            this.checkAddinProjectDescription.Name = "checkAddinProjectDescription";
            this.checkAddinProjectDescription.Size = new System.Drawing.Size(538, 37);
            this.checkAddinProjectDescription.TabIndex = 1;
            this.checkAddinProjectDescription.Text = "Adds project for the add-in. Nte taht you can skip addin project creation and cre" +
    "ate only installer project to install Visio files. ";
            // 
            // radioAddinTypeVSTOLabel
            // 
            this.radioAddinTypeVSTOLabel.ForeColor = System.Drawing.SystemColors.GrayText;
            this.radioAddinTypeVSTOLabel.LinkArea = new System.Windows.Forms.LinkArea(229, 47);
            this.radioAddinTypeVSTOLabel.LinkColor = System.Drawing.Color.Sienna;
            this.radioAddinTypeVSTOLabel.Location = new System.Drawing.Point(78, 393);
            this.radioAddinTypeVSTOLabel.Name = "radioAddinTypeVSTOLabel";
            this.radioAddinTypeVSTOLabel.Size = new System.Drawing.Size(516, 42);
            this.radioAddinTypeVSTOLabel.TabIndex = 10;
            this.radioAddinTypeVSTOLabel.Text = "Creates the Addin project using Visual Studio Tools for Office. Select this optio" +
    "n for visual studio designer support.";
            this.radioAddinTypeVSTOLabel.UseCompatibleTextRendering = true;
            this.radioAddinTypeVSTOLabel.VisitedLinkColor = System.Drawing.SystemColors.HotTrack;
            this.radioAddinTypeVSTOLabel.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.checkAddinProjectDescription_LinkClicked);
            // 
            // addinTypeLabel
            // 
            this.addinTypeLabel.AutoSize = true;
            this.addinTypeLabel.Location = new System.Drawing.Point(40, 345);
            this.addinTypeLabel.Name = "addinTypeLabel";
            this.addinTypeLabel.Size = new System.Drawing.Size(90, 13);
            this.addinTypeLabel.TabIndex = 8;
            this.addinTypeLabel.Text = "Type of the addin";
            // 
            // addinDescriptionLabel
            // 
            this.addinDescriptionLabel.AutoSize = true;
            this.addinDescriptionLabel.Location = new System.Drawing.Point(40, 253);
            this.addinDescriptionLabel.Name = "addinDescriptionLabel";
            this.addinDescriptionLabel.Size = new System.Drawing.Size(88, 13);
            this.addinDescriptionLabel.TabIndex = 6;
            this.addinDescriptionLabel.Text = "Addin description";
            // 
            // addinManufacturerLabel
            // 
            this.addinManufacturerLabel.AutoSize = true;
            this.addinManufacturerLabel.Location = new System.Drawing.Point(37, 198);
            this.addinManufacturerLabel.Name = "addinManufacturerLabel";
            this.addinManufacturerLabel.Size = new System.Drawing.Size(99, 13);
            this.addinManufacturerLabel.TabIndex = 4;
            this.addinManufacturerLabel.Text = "Addin manufacturer";
            // 
            // addinNameLabel
            // 
            this.addinNameLabel.AutoSize = true;
            this.addinNameLabel.Location = new System.Drawing.Point(37, 148);
            this.addinNameLabel.Name = "addinNameLabel";
            this.addinNameLabel.Size = new System.Drawing.Size(63, 13);
            this.addinNameLabel.TabIndex = 2;
            this.addinNameLabel.Text = "Addin name";
            // 
            // addinManufacturer
            // 
            this.addinManufacturer.Location = new System.Drawing.Point(37, 216);
            this.addinManufacturer.Name = "addinManufacturer";
            this.addinManufacturer.Size = new System.Drawing.Size(385, 20);
            this.addinManufacturer.TabIndex = 5;
            // 
            // addinDescription
            // 
            this.addinDescription.Location = new System.Drawing.Point(37, 271);
            this.addinDescription.Multiline = true;
            this.addinDescription.Name = "addinDescription";
            this.addinDescription.Size = new System.Drawing.Size(385, 50);
            this.addinDescription.TabIndex = 7;
            // 
            // addinName
            // 
            this.addinName.Location = new System.Drawing.Point(37, 165);
            this.addinName.Name = "addinName";
            this.addinName.Size = new System.Drawing.Size(385, 20);
            this.addinName.TabIndex = 3;
            // 
            // radioAddinTypeCOMLabel
            // 
            this.radioAddinTypeCOMLabel.ForeColor = System.Drawing.SystemColors.GrayText;
            this.radioAddinTypeCOMLabel.Location = new System.Drawing.Point(75, 461);
            this.radioAddinTypeCOMLabel.Name = "radioAddinTypeCOMLabel";
            this.radioAddinTypeCOMLabel.Size = new System.Drawing.Size(401, 23);
            this.radioAddinTypeCOMLabel.TabIndex = 12;
            this.radioAddinTypeCOMLabel.Text = "Use this option to create COM addin. Does not use VSTO.";
            // 
            // radioAddinTypeCOM
            // 
            this.radioAddinTypeCOM.AutoSize = true;
            this.radioAddinTypeCOM.Location = new System.Drawing.Point(56, 438);
            this.radioAddinTypeCOM.Name = "radioAddinTypeCOM";
            this.radioAddinTypeCOM.Size = new System.Drawing.Size(157, 17);
            this.radioAddinTypeCOM.TabIndex = 11;
            this.radioAddinTypeCOM.Text = "Create a COM Addin project";
            this.radioAddinTypeCOM.UseVisualStyleBackColor = true;
            // 
            // radioAddinTypeVSTO
            // 
            this.radioAddinTypeVSTO.AutoSize = true;
            this.radioAddinTypeVSTO.Checked = true;
            this.radioAddinTypeVSTO.Location = new System.Drawing.Point(56, 373);
            this.radioAddinTypeVSTO.Name = "radioAddinTypeVSTO";
            this.radioAddinTypeVSTO.Size = new System.Drawing.Size(162, 17);
            this.radioAddinTypeVSTO.TabIndex = 9;
            this.radioAddinTypeVSTO.TabStop = true;
            this.radioAddinTypeVSTO.Text = "Create a VSTO Addin project";
            this.radioAddinTypeVSTO.UseVisualStyleBackColor = true;
            // 
            // checkAddinProject
            // 
            this.checkAddinProject.AutoSize = true;
            this.checkAddinProject.Location = new System.Drawing.Point(24, 88);
            this.checkAddinProject.Name = "checkAddinProject";
            this.checkAddinProject.Size = new System.Drawing.Size(147, 17);
            this.checkAddinProject.TabIndex = 0;
            this.checkAddinProject.Text = "Create Visio Addin project";
            this.checkAddinProject.UseVisualStyleBackColor = true;
            this.checkAddinProject.CheckedChanged += new System.EventHandler(this.UpdateButtons);
            // 
            // WizardForm
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            this.ClientSize = new System.Drawing.Size(697, 554);
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
            this.pageAddinOptions.ResumeLayout(false);
            this.pageAddinOptions.PerformLayout();
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
            //if (e.OldIndex == 0 && !checkAddinProject.Checked)
            //    e.NewIndex = 2;

            //if (e.OldIndex == 2 && !checkAddinProject.Checked)
            //    e.NewIndex = 0;
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

            var enableAddinVSTO = checkAddinProject.Checked && _host.IsVstoInstalled();
            radioAddinTypeCOM.Enabled = checkAddinProject.Checked;
            radioAddinTypeCOMLabel.ForeColor = checkAddinProject.Checked ? SystemColors.GrayText : SystemColors.ControlDark;

            radioAddinTypeVSTO.Enabled = enableAddinVSTO;
            if (_host.IsVstoInstalled())
                radioAddinTypeVSTOLabel.ForeColor = radioAddinTypeVSTO.Enabled ? SystemColors.GrayText : SystemColors.ControlDark;

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
            checkEnableSetupUIDescription.LinkColor = checkWixSetup.Checked ? SystemColors.HotTrack : Color.CornflowerBlue;


            checkEnableSetupLanguage.Enabled = checkWixSetup.Checked;
            checkEnableSetupLanguageDescription.ForeColor = checkWixSetup.Checked ? SystemColors.GrayText : SystemColors.ControlDark;
            checkEnableSetupLanguageDescription.LinkColor = checkWixSetup.Checked ? SystemColors.HotTrack : Color.CornflowerBlue;

            comboSetupLanguage.Enabled = checkWixSetup.Checked && checkEnableSetupLanguage.Checked;

            var isWixInstalled = _host.IsWixInstalled();
            if (isWixInstalled)
                checkWixSetupDescription.ForeColor = checkWixSetup.Checked ? SystemColors.GrayText : SystemColors.ControlDark;

            comboSetupUI.Enabled = checkWixSetup.Checked && checkEnableSetupUI.Checked;

            var visioInteropVersion = GetVisioInteropVersion();
            var enableLocalReferences = checkAddinProject.Checked && visioInteropVersion > 0;
            checkLocalReferences.Enabled = enableLocalReferences;
            if (visioInteropVersion > 0)
                checkLocalReferencesLabel.ForeColor = enableLocalReferences ? SystemColors.GrayText : SystemColors.ControlDark;

            if (radioAddinTypeCOM.Checked)
            {
                radioSupportRibbonDesigner.Checked = false;
                radioSupportRibbonXml.Checked = true;
            }

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
            OpenExternalLink(ExternalLink.WixDocsUI);
        }

        private void checkAddinProjectDescription_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            OpenExternalLink(ExternalLink.VstoDownload);
        }

        private void checkEnableSetupLanguageDescription_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            OpenExternalLink(ExternalLink.WixDocLanguage);
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
