using System.Windows.Forms;
using System.Collections.Generic;
using EnvDTE;
using Microsoft.VisualStudio.TemplateWizard;
using Microsoft.Win32;

namespace PanelAddinWizard
{
	/// <summary>
	/// Summary description for SampleWizard.
	/// </summary>
	public class RootWizard 
        : IWizard
	{
        public static Dictionary<string, string> GlobalDictionary =
            new Dictionary<string, string>();

	    private DTE _dte;

        // Add global replacement parameters
        public void RunStarted(object automationObject,
            Dictionary<string, string> replacementsDictionary,
            WizardRunKind runKind, object[] customParams)
        {
            _dte = automationObject as DTE;

            if (_dte != null)
            {
                var vstoKey = string.Format(@"HKEY_LOCAL_MACHINE\{0}\Setup\VSTO", _dte.RegistryRoot);
                var val = Registry.GetValue(vstoKey, "ProductDir", null);
                
                if (null == val)
                {
                    MessageBox.Show(
                        "Visual Studio Tools for Office installation is not detected. " +
                        "This project template is based on VSTO project, thus it is required. " +
                        "Please install Tools for Office, or consder using COM Visio Addin project template.",
                        "Visio VSTO Panel Addin: VSTO not found", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                    throw new WizardBackoutException();
                }
            }

            var wizardForm = new WizardForm
            {
                TaskPane = true, 
                Ribbon = true
            };

            if (wizardForm.ShowDialog() == DialogResult.Cancel)
                throw new WizardBackoutException();

            replacementsDictionary["$csprojectname$"] = replacementsDictionary["$safeprojectname$"];
            replacementsDictionary["$progid$"] = replacementsDictionary["$safeprojectname$"] + ".Addin";
            replacementsDictionary["$clsid$"] = replacementsDictionary["$guid1$"];
            replacementsDictionary["$csprojectguid$"] = replacementsDictionary["$guid2$"];
            replacementsDictionary["$mergeguid$"] = replacementsDictionary["$guid4$"];

            replacementsDictionary["$ribbon$"] = wizardForm.Ribbon ? "true" : "false";
            replacementsDictionary["$commandbars$"] = wizardForm.CommandBars ? "true" : "false";

            replacementsDictionary["$ribbonORcommandbars$"] = wizardForm.Ribbon || wizardForm.CommandBars ? "true" : "false";
            replacementsDictionary["$ribbonANDcommandbars$"] = wizardForm.Ribbon && wizardForm.CommandBars ? "true" : "false";
            replacementsDictionary["$commandbarsANDtaskpane$"] = wizardForm.CommandBars && wizardForm.TaskPane ? "true" : "false";
            replacementsDictionary["$taskpane$"] = wizardForm.TaskPane ? "true" : "false";
            replacementsDictionary["$ui$"] = (wizardForm.CommandBars || wizardForm.Ribbon) ? "true" : "false";
            replacementsDictionary["$taskpaneANDui$"] = (wizardForm.TaskPane && (wizardForm.CommandBars || wizardForm.Ribbon)) ? "true" : "false";
            replacementsDictionary["$taskpaneORui$"] = (wizardForm.TaskPane || (wizardForm.CommandBars || wizardForm.Ribbon)) ? "true" : "false";

            replacementsDictionary["$office$"] = GetOfficeVersion();
        }

        // Don't return the latest version, even if it is installed, because it will cause problems by auto-upgrade
	    static string GetOfficeVersion()
	    {
            return Registry.GetValue(@"HKEY_LOCAL_MACHINE\Software\Microsoft\Office\14.0\Visio\InstallRoot", "Path", null) != null
                ? "14.0" : "12.0";
	    }

        public void RunFinished()
        {
        }

        public void BeforeOpeningFile(ProjectItem projectItem)
        {
        }

        public void ProjectFinishedGenerating(Project project)
        {
        }

        public bool ShouldAddProjectItem(string filePath)
        {
            return true;
        }

        public void ProjectItemFinishedGenerating(ProjectItem projectItem)
        {
        }
    }
}
