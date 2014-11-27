$if$ ($ui$ == true)using System.Drawing;
$endif$$if$ ($ribbonANDcommandbars$ == true)using System.Globalization;
$endif$$if$ ($ui$ == true)using System.Windows.Forms;
$endif$$if$ ($ui$ == true)using $csprojectname$.Properties;
$endif$using Visio = Microsoft.Office.Interop.Visio;

namespace $csprojectname$
{
    public$if$ ($ui$ == true) partial$endif$ class Addin
    {
        public Visio.Application Application { get; set; }
        $if$ ($ui$ == true)
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

                case "Command2":    // make command2 enabled only if a shape is selected
                    return Application != null && Application.ActiveWindow != null && Application.ActiveWindow.Selection.Count > 0;
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
        #region Panel
        private void TogglePanel()
        {
            _panelManager.TogglePanel(Application.ActiveWindow);
        }

        private bool IsPanelEnabled()
        {
            return Application != null && Application.ActiveWindow != null;
        }

        private bool IsPanelVisible()
        {
            return Application != null && _panelManager.IsPanelOpened(Application.ActiveWindow);
        }
        
        private PanelManager _panelManager;
        #endregion
        $endif$
        internal void Startup(object application)
        {
            Application = (Visio.Application) application;
            $if$ ($taskpane$ == true)
            _panelManager = new PanelManager(this);
            $endif$$if$ ($ribbonANDcommandbars$ == true)var version = int.Parse(Application.Version, NumberStyles.AllowDecimalPoint);
            if (version < 14)
                $endif$$if$ ($commandbars$ == true)StartupCommandBars("$csprojectname$", new[] { "Command1", "Command2" $endif$$if$ ($commandbarsANDtaskpane$ == true) , "TogglePanel"$endif$$if$ ($commandbars$ == true)});
        $endif$}

        internal void Shutdown()
        {
            $if$ ($commandbars$ == true)ShutdownCommandBars();
            $endif$$if$ ($taskpane$ == true)_panelManager.Dispose();
        $endif$}
        $if$ ($ui$ == true)
        internal void UpdateUI()
        {
            $endif$$if$ ($commandbars$ == true) UpdateCommandBars();
            $endif$$if$ ($ribbon$ == true)UpdateRibbon();
        $endif$$if$ ($ui$ == true)}
    $endif$}
}
