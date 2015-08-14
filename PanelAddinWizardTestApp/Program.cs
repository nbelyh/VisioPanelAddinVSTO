using System;
using System.Text.RegularExpressions;
using PanelAddinWizard;

namespace PanelAddinWizardTestApp
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

        public void OpenExternalLink(ExternalLink link)
        {
            var url = GetExternalLinkUrl(link);

            if (url != null)
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(url));
        }

        string GetExternalLinkUrl(ExternalLink link)
        {
            switch (link)
            {
                case ExternalLink.WixDocsUI:
                    return "http://wixtoolset.org/documentation/manual/v3/wixui/wixui_dialog_library.html";

                case ExternalLink.WixDownload:
                    return "http://wixtoolset.org/";

                case ExternalLink.VstoDownload:
                    {
                        var match = Regex.Match("12.0", @"\d+");
                        if (match.Success)
                        {
                            if (match.Value == "11")
                                return "http://aka.ms/OfficeDevToolsForVS2012";
                            if (match.Value == "12")
                                return "http://aka.ms/OfficeDevToolsForVS2013";
                            if (match.Value == "14")
                                return "http://aka.ms/OfficeDevToolsForVS2015";
                        }
                        return "http://dev.office.com/";
                    }
            }

            return null;
        }
    }

    class Program
    {
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
