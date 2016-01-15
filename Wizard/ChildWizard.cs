using System;
using System.Collections.Generic;
using System.IO;
using EnvDTE;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.TemplateWizard;

namespace PanelAddinWizard
{
    // Child wizard is used by child project vstemplates

    public class ChildWizard : IWizard
    {
        // Retrieve global replacement parameters
        public void RunStarted(object automationObject, 
            Dictionary<string, string> replacementsDictionary, 
            WizardRunKind runKind, object[] customParams)
        {
            if (replacementsDictionary["$safeprojectname$"] == "Setup")
            {
                if (RootWizard.GlobalDictionary["$wixSetup$"] != "true")
                    throw new WizardBackoutException("Cancel generation of wix setup project");
            }
            else if (RootWizard.GlobalDictionary["$addinProject$"] != "true")
            {
                throw new WizardBackoutException("Cancel generation of addin project");
            }

            // Add custom parameters.
            foreach (var kvp in RootWizard.GlobalDictionary)
                replacementsDictionary.Add(kvp.Key, kvp.Value);
        }

        public void RunFinished()
        {
        }

        public void BeforeOpeningFile(ProjectItem projectItem)
        {
        }

        public void ProjectFinishedGenerating(Project project)
        {
            if (Path.GetExtension(project.FileName) == ".csproj" || Path.GetExtension(project.FileName) == ".vbproj")
            {
                RootWizard.GlobalDictionary["$csprojectguid$"] = GetProjectGuid(project);
                SetStartupAction(project);
            }

            if (Path.GetExtension(project.FileName) == ".wixproj")
            {
                var projectPath = Path.GetDirectoryName(project.FullName);

                GenerateVisioFiles(projectPath);
            }
        }

        void GenerateVisioFiles(string projectPath)
        {
            if (!RootWizard.HaveVisioFiles)
                return;

            if (!RootWizard.WizardOptions.DuplicateExistingVisioFiles)
                return;

            foreach (var path in RootWizard.WizardOptions.VisioFilePaths)
                File.Copy(path, Path.Combine(projectPath, Path.GetFileName(path)));
        }

        private void SetStartupAction(Project project)
        {
            try
            {
                var path32 = RootWizard.GetVisioPath32();
                var path64 = RootWizard.GetVisioPath64();

                var configManager = project.ConfigurationManager;
                for (var i = 1; i <= configManager.Count; ++i)
                {
                    var config = configManager.Item(i);
                    if (path32 != null)
                    {
                        config.Properties.Item("StartAction").Value = 1;
                        config.Properties.Item("StartProgram").Value = path32;
                    }
                    if (path64 != null)
                    {
                        config.Properties.Item("StartAction").Value = 1;
                        config.Properties.Item("StartProgram").Value = path64;
                    }
                }
            }
            // this convenience feature; continue if failed, not a big deal
            // ReSharper disable once EmptyGeneralCatchClause
            catch (Exception)
            {
            }
        }

        public static string GetProjectGuid(Project project)
        {
            IServiceProvider serviceProvider = new ServiceProvider(project.DTE as
                Microsoft.VisualStudio.OLE.Interop.IServiceProvider);

            var sol = (IVsSolution)serviceProvider.GetService(typeof(SVsSolution));

            IVsHierarchy proj;

            ErrorHandler.ThrowOnFailure(sol.GetProjectOfUniqueName(project.UniqueName, out proj));

            Guid projGuid;

            proj.GetGuidProperty(VSConstants.VSITEMID_ROOT, (int)__VSHPROPID.VSHPROPID_ProjectIDGuid, out projGuid);

            return projGuid.ToString();
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
