
using System.Drawing;
using PanelAddinWizard.Properties;

namespace PanelAddinWizard
{
    public class RootWizardCS : RootWizard
    {
        protected override Image HeaderImage
        {
            get { return Resources.IconCS.ToBitmap(); }
        }
    }
}
