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

    $if$ ($vstoAddin$ == true)[ComVisible(true)]
    $endif$public partial class $if$ ($vstoAddin$ == true)AddinUI$else$ThisAddIn$endif$ : Office.IRibbonExtensibility
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
            return $thisAddIn$IsCommandEnabled(ctrl.Id);
        }

        public bool IsRibbonCommandChecked(Office.IRibbonControl ctrl)
        {
            return $thisAddIn$IsCommandChecked(ctrl.Id);
        }

        public void OnRibbonButtonCheckClick(Office.IRibbonControl control, bool pressed)
        {
            $thisAddIn$OnCommand(control.Id);
        }

        public void OnRibbonButtonClick(Office.IRibbonControl control)
        {
            $thisAddIn$OnCommand(control.Id);
        }

        public string OnGetRibbonLabel(Office.IRibbonControl control)
        {
            return $thisAddIn$GetCommandLabel(control.Id);
        }

        public void OnRibbonLoad(Office.IRibbonUI ribbonUI)
        {
            _ribbon = ribbonUI;
        }

        public Bitmap GetRibbonImage(Office.IRibbonControl control)
        {
            return $thisAddIn$GetCommandBitmap(control.Id);
        }

        #endregion

        public void UpdateRibbon()
        {
            _ribbon.Invalidate();
        }
    }
}
