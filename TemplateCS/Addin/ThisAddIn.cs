using System;
using Office = Microsoft.Office.Core;
$if$ ($ui$ == true)using System.Drawing;
$endif$$if$ ($ribbonANDcommandbars$ == true)using System.Globalization;
$endif$$if$ ($ui$ == true)using System.Windows.Forms;
$endif$$if$ ($ui$ == true)using $csprojectname$.Properties;
$endif$using Visio = Microsoft.Office.Interop.Visio;

namespace $csprojectname$
{
    public partial class ThisAddIn
    {
        $if$ ($commandbars$ == true)
        private readonly AddinCommandBars _addinCommandBars = new AddinCommandBars();
        $endif$$if$ ($ribbon$ == true)
        private readonly AddinRibbon _addinRibbon = new AddinRibbon();
        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return _addinRibbon;
        }
        $endif$$if$ ($ui$ == true)
        /// <summary>
        /// Callback called by the UI manager when user clicks a button
        /// Should do something meaninful wehn corresponding action is called.
        /// </summary>
        public void OnCommand(string commandId)
        {
            switch (commandId)
            {
                case "Command1":
                    MessageBox.Show(commandId);
                    return;

                case "Command2":
                    MessageBox.Show(commandId);
                    return;
                $endif$$if$ ($taskpaneANDui$ == true)
                case "TogglePanel":
                    TogglePanel();
                    return;
            $endif$$if$ ($ui$ == true)}
        }

        /// <summary>
        /// Callback called by UI manager.
        /// Should return if corresponding command shoudl be enabled in the user interface.
        /// By default, all commands are enabled.
        /// </summary>
        public bool IsCommandEnabled(string commandId)
        {
            switch (commandId)
            {
                case "Command1":    // make command1 always enabled
                    return true;

                case "Command2":    // make command2 enabled only if a drawing is opened
                    return Application != null && Application.ActiveWindow != null;
                $endif$$if$ ($taskpaneANDui$ == true)
                case "TogglePanel": // make panel enabled only if we have an open drawing.
                    return IsPanelEnabled();

                $endif$$if$ ($ui$ == true) default:
                    return true;
            }
        }

        /// <summary>
        /// Callback called by UI manager.
        /// Should return if corresponding command (button) is pressed or not (makes sense for toggle buttons)
        /// </summary>
        public bool IsCommandChecked(string command)
        {
            $endif$$if$ ($taskpaneANDui$ == true)
            if (command == "TogglePanel")
                return IsPanelVisible();
            
            $endif$$if$($ui$ == true)return false;
        }
        /// <summary>
        /// Callback called by UI manager.
        /// Returns a label associated with given command.
        /// We assume for simplicity taht command labels are named simply named as [commandId]_Label (see resources)
        /// </summary>
        public string GetCommandLabel(string command)
        {
            return Resources.ResourceManager.GetString(command + "_Label");
        }

        /// <summary>
        /// Returns a icon associated with given command.
        /// We assume for simplicity that icon ids are named after command commandId.
        /// </summary>
        public Icon GetCommandIcon(string command)
        {
            return (Icon)Resources.ResourceManager.GetObject(command);
        }
        $endif$$if$ ($taskpane$ == true)
        public void TogglePanel()
        {
            _panelManager.TogglePanel(Application.ActiveWindow);
        }

        public bool IsPanelEnabled()
        {
            return Application != null && Application.ActiveWindow != null;
        }

        public bool IsPanelVisible()
        {
            return Application != null && _panelManager.IsPanelOpened(Application.ActiveWindow);
        }
        
        private PanelManager _panelManager;
        $endif$
        $if$ ($taskpaneORui$ == true)
        internal void UpdateUI()
        {
            $endif$$if$ ($commandbars$ == true) _addinCommandBars.UpdateCommandBars();
            $endif$$if$ ($ribbon$ == true)_addinRibbon.UpdateRibbon();
        $endif$$if$ ($taskpaneORui$ == true)}
        $endif$
        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            $if$ ($taskpane$ == true)_panelManager = new PanelManager();
            $endif$$if$ ($ribbonANDcommandbars$ == true)var version = int.Parse(Application.Version, NumberStyles.AllowDecimalPoint);
            if (version < 14)
                $endif$$if$ ($commandbars$ == true)_addinCommandBars.StartupCommandBars("$csprojectname$", new[] { "Command1", "Command2" $endif$$if$ ($commandbarsANDtaskpane$ == true) , "TogglePanel"$endif$$if$ ($commandbars$ == true)});
        $endif$}

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
        $if$ ($commandbars$ == true)_addinCommandBars.ShutdownCommandBars();
            $endif$$if$ ($taskpane$ == true)_panelManager.Dispose();
        $endif$}
        
        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            Startup += ThisAddIn_Startup;
            Shutdown += ThisAddIn_Shutdown;
        }

        #endregion
    }
}
