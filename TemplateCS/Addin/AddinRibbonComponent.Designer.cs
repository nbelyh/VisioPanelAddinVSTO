namespace $csprojectname$
{
    partial class AddinRibbonComponent : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public AddinRibbonComponent()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

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

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.Tab1 = this.Factory.CreateRibbonTab();
            this.Group1 = this.Factory.CreateRibbonGroup();
            this.Command1 = this.Factory.CreateRibbonButton();
            $if$ ($taskpane$ == true)this.TogglePanel = this.Factory.CreateRibbonButton();
            $endif$this.Tab1.SuspendLayout();
            this.Group1.SuspendLayout();
            // 
            // Tab1
            // 
            this.Tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.Tab1.ControlId.OfficeId = "TabHome";
            this.Tab1.Groups.Add(this.Group1);
            this.Tab1.Label = "TabHome";
            this.Tab1.Name = "Tab1";
            // 
            // Group1
            // 
            this.Group1.Items.Add(this.Command1);
            $if$ ($taskpane$ == true)this.Group1.Items.Add(this.TogglePanel);
            $endif$this.Group1.Label = "Group1";
            this.Group1.Name = "Group1";
            // 
            // Command1
            // 
            this.Command1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.Command1.Image = global::$csprojectname$.Properties.Resources.Command1;
            this.Command1.Label = "Command 1";
            this.Command1.Name = "Command1";
            this.Command1.ShowImage = true;
            this.Command1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonCommand1_Click);
            $if$ ($taskpane$ == true)
            // 
            // TogglePanel
            // 
            this.TogglePanel.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.TogglePanel.Image = global::$csprojectname$.Properties.Resources.TogglePanel;
            this.TogglePanel.Label = "Toggle Panel";
            this.TogglePanel.Name = "TogglePanel";
            this.TogglePanel.ShowImage = true;
            this.TogglePanel.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonToggle_Click);
            $endif$
            // 
            // AddinRibbon
            // 
            this.Name = "AddinRibbon";
            this.RibbonType = "Microsoft.Visio.Drawing";
            this.Tabs.Add(this.Tab1);
            this.Tab1.ResumeLayout(false);
            this.Tab1.PerformLayout();
            this.Group1.ResumeLayout(false);
            this.Group1.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab Tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup Group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Command1;
        $if$ ($taskpane$ == true)internal Microsoft.Office.Tools.Ribbon.RibbonButton TogglePanel;
    $endif$}

    partial class ThisRibbonCollection
    {
        internal AddinRibbonComponent Ribbon
        {
            get { return this.GetRibbon<AddinRibbonComponent>(); }
        }
    }
}
