using System.Drawing;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;

namespace $csprojectname$
{
    /// <summary>
    /// User interface manager for Visio 2010 and above
    /// Creates and controls ribbon UI
    /// </summary>
    /// 
    [ComVisible(true)]
    public partial class Addin
        : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI _ribbon;

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonId)
        {
            return Properties.Resources.Ribbon;
        }

        #endregion

        #region Ribbon Callbacks

        public bool IsRibbonCommandEnabled(Office.IRibbonControl ctrl)
        {
            return IsCommandEnabled(ctrl.Id);
        }

        public bool IsRibbonCommandChecked(Office.IRibbonControl ctrl)
        {
            return IsCommandChecked(ctrl.Id);
        }

        public void OnRibbonButtonCheckClick(Office.IRibbonControl control, bool pressed)
        {
            OnCommand(control.Id);
        }

        public void OnRibbonButtonClick(Office.IRibbonControl control)
        {
            OnCommand(control.Id);
        }

        public string OnGetRibbonLabel(Office.IRibbonControl control)
        {
            return GetCommandLabel(control.Id);
        }

        public void OnRibbonLoad(Office.IRibbonUI ribbonUI)
        {
            _ribbon = ribbonUI;
        }

        public Bitmap GetRibbonImage(Office.IRibbonControl control)
        {
            var icon = GetCommandIcon(control.Id);
            return icon != null ? icon.ToBitmap() : null;
        }

        #endregion

        public void UpdateRibbon()
        {
            _ribbon.Invalidate();
        }
    }
}
