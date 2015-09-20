using System;
using System.Collections.Generic;
using Microsoft.Office.Tools.Ribbon;

namespace $csprojectname$
{
    public partial class AddinRibbonComponent
    {
        private void buttonCommand1_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Command1();
        }
        $if$ ($taskpane$ == true)
        private void buttonToggle_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.TogglePanel();
        }
    $endif$}
}
