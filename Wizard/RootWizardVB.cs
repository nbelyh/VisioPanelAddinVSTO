
using System.Drawing;
using PanelAddinWizard.Properties;

namespace PanelAddinWizard
{
    public class RootWizardVB : RootWizard
    {
        protected override Image HeaderImage
        {
            get { return Resources.IconVB.ToBitmap(); }
        }
    }
}
