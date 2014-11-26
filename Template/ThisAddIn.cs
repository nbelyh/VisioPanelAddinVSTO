using System;
using Office = Microsoft.Office.Core;

namespace $csprojectname$
{
    public partial class ThisAddIn
    {
        private readonly Addin _addin = new Addin();
        $if$ ($ribbon$ == true)
        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return _addin;
        }
        $endif$
        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            _addin.Startup(Application);
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            _addin.Shutdown();
        }
        
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
