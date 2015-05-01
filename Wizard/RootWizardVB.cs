
using System.Drawing;
using PanelAddinWizard.Properties;

namespace PanelAddinWizard
{
    /// <summary>
    /// Summary description for WizardForm.
    /// </summary>
    public class RootWizardVB : RootWizard
    {
        protected override Image HeaderImage
        {
            get { return Resources.IconVB.ToBitmap(); }
        }
    }
}
