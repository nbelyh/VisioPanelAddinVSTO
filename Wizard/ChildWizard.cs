﻿using System;
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
            // Add custom parameters.
            replacementsDictionary.Add("$csprojectname$", RootWizard.GlobalDictionary["$csprojectname$"]);
            replacementsDictionary.Add("$progid$", RootWizard.GlobalDictionary["$progid$"]);
            replacementsDictionary.Add("$clsid$", RootWizard.GlobalDictionary["$clsid$"]);
            replacementsDictionary.Add("$csprojectguid$", RootWizard.GlobalDictionary["$csprojectguid$"]);
            replacementsDictionary.Add("$wixproject$", RootWizard.GlobalDictionary["$wixproject$"]);

            replacementsDictionary.Add("$mergeguid$", RootWizard.GlobalDictionary["$mergeguid$"]);

            replacementsDictionary.Add("$ribbon$", RootWizard.GlobalDictionary["$ribbon$"]);
            replacementsDictionary.Add("$commandbars$", RootWizard.GlobalDictionary["$commandbars$"]);

            replacementsDictionary.Add("$ribbonORcommandbars$", RootWizard.GlobalDictionary["$ribbonORcommandbars$"]);
            replacementsDictionary.Add("$ribbonANDcommandbars$", RootWizard.GlobalDictionary["$ribbonANDcommandbars$"]);
            replacementsDictionary.Add("$commandbarsANDtaskpane$", RootWizard.GlobalDictionary["$commandbarsANDtaskpane$"]);
            replacementsDictionary.Add("$taskpane$", RootWizard.GlobalDictionary["$taskpane$"]);
            replacementsDictionary.Add("$ui$", RootWizard.GlobalDictionary["$ui$"]);
            replacementsDictionary.Add("$taskpaneANDui$", RootWizard.GlobalDictionary["$taskpaneANDui$"]);
            replacementsDictionary.Add("$taskpaneORui$", RootWizard.GlobalDictionary["$taskpaneORui$"]);

            replacementsDictionary.Add("$office$", RootWizard.GlobalDictionary["$office$"]);
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