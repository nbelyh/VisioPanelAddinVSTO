using System;
using PanelAddinWizard;

namespace PanelAddinWizardTestApp
{
    class Program
    {
        class TestHost : IWizardFormHost
        {
            public bool IsWixInstalled()
            {
                return true;
            }

            public bool IsVstoInstalled()
            {
                return true;
            }

            public int GetVisualStudioVersion()
            {
                return 12;
            }
        }

        [STAThread]
        static void Main(string[] args)
        {
            var wizardForm = new WizardForm(new TestHost())
            {
            };

            wizardForm.ShowDialog();

            wizardForm.GetSetupOptions();
            
            //var f = GetFiles("visio:PublishStencil", wizardForm.StencilPaths);
            //var g = GetFiles("visio:PublishTemplate", wizardForm.TemplatePaths);
        }
    }
}
