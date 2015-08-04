using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using PanelAddinWizard;

namespace PanelAddinWizardTestApp
{
    
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            var wizardForm = new WizardForm(true)
            {
            };

            wizardForm.ShowDialog();

            var s = wizardForm.StencilPaths;
            var t = wizardForm.TemplatePaths;
        }
    }
}
