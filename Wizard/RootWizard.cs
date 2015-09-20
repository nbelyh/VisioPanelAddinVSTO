
using System;
using System.IO;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Drawing;
using System.Xml;
using EnvDTE;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.TemplateWizard;
using Microsoft.Win32;
using System.Text;
using EnvDTE80;
using System.Text.RegularExpressions;

namespace PanelAddinWizard
{
    /// <summary>
    /// Summary description for WizardForm.
    /// </summary>
    public abstract class RootWizard
        : IWizard
        , IWizardFormHost
    {
        // Use to communicate $saferootprojectname$ to ChildWizard
        public static Dictionary<string, string> GlobalDictionary =
            new Dictionary<string, string>();

        public static WixSetupOptions SetupOptions;

        private DTE2 _dte;

        protected abstract Image HeaderImage { get; }

        // Add global replacement parameters
        public void RunStarted(object automationObject,
            Dictionary<string, string> replacementsDictionary,
            WizardRunKind runKind, object[] customParams)
        {
            _dte = automationObject as DTE2;

            var wizardForm = new WizardForm(this)
            {
                HeaderImage = HeaderImage
            };

            if (wizardForm.ShowDialog() == DialogResult.Cancel)
                throw new WizardBackoutException();

            GlobalDictionary["$csprojectname$"] = replacementsDictionary["$safeprojectname$"];
            GlobalDictionary["$progid$"] = replacementsDictionary["$safeprojectname$"] + ".Addin";
            GlobalDictionary["$clsid$"] = replacementsDictionary["$guid1$"];
            GlobalDictionary["$wixproject$"] = replacementsDictionary["$guid2$"];
            GlobalDictionary["$csprojectguid$"] = replacementsDictionary["$guid3$"];

            GlobalDictionary["$mergeguid$"] = replacementsDictionary["$guid4$"];

            var ribbon = wizardForm.RibbonXml || wizardForm.RibbonComponent;

            GlobalDictionary["$ribbon$"] = ribbon ? "true" : "false";
            GlobalDictionary["$ribbonXml$"] = wizardForm.RibbonXml ? "true" : "false";
            GlobalDictionary["$ribbonComponent$"] = wizardForm.RibbonComponent ? "true" : "false";

            GlobalDictionary["$commandbars$"] = wizardForm.CommandBars ? "true" : "false";

            GlobalDictionary["$ribbonANDcommandbars$"] = ribbon  && wizardForm.CommandBars ? "true" : "false";
            GlobalDictionary["$commandbarsANDtaskpane$"] = wizardForm.CommandBars && wizardForm.TaskPane ? "true" : "false";
            GlobalDictionary["$taskpane$"] = wizardForm.TaskPane ? "true" : "false";

            var ui = (wizardForm.CommandBars || ribbon);
            GlobalDictionary["$ui$"] = ui ? "true" : "false";
            GlobalDictionary["$taskpaneANDui$"] = (wizardForm.TaskPane && ui) ? "true" : "false";
            GlobalDictionary["$taskpaneORui$"] = (wizardForm.TaskPane || ui) ? "true" : "false";

            var uiCallbacks = (wizardForm.CommandBars || wizardForm.RibbonXml);
            GlobalDictionary["$uiCallbacks$"] = uiCallbacks ? "true" : "false";
            GlobalDictionary["$taskpaneANDuiCallbacks$"] = (wizardForm.TaskPane && uiCallbacks) ? "true" : "false";

            GlobalDictionary["$addinProject$"] = wizardForm.AddinEnabled ? "true" : "false";
            GlobalDictionary["$office$"] = GetOfficeVersion();
            GlobalDictionary["$wixSetup$"] = wizardForm.WixSetup ? "true" : "false";

            SetupOptions = wizardForm.GetSetupOptions();
            GetFiles(SetupOptions);

            GlobalDictionary["$EnableWixUI$"] = SetupOptions.EnableWixUI ? "true" : "false";
            GlobalDictionary["$WixUI$"] = SetupOptions.EnableWixUI ? SetupOptions.WixUI : "";
            GlobalDictionary["$defaultVisioFiles$"] = SetupOptions.EnableWixSetup && SetupOptions.CreateNewVisioFiles ? "true" : "false";

            GlobalDictionary["$EnableLicense$"] = SetupOptions.EnableLicense ? "true" : "false";
            GlobalDictionary["$LicenseFileName$"] = Path.GetFileName(SetupOptions.LicenseFilePath);
        }

        private static string beautifyXml(XmlDocument doc)
        {
            var sb = new StringBuilder();
            var settings =
                new XmlWriterSettings
                {
                    Indent = true,
                    IndentChars = @"    ",
                    NewLineChars = Environment.NewLine,
                    NewLineHandling = NewLineHandling.Replace,
                    OmitXmlDeclaration = true
                };

            using (var writer = XmlWriter.Create(sb, settings))
            {
                if (doc.ChildNodes[0] is XmlProcessingInstruction)
                {
                    doc.RemoveChild(doc.ChildNodes[0]);
                }

                doc.Save(writer);
                return sb.ToString();
            }
        }

        void GetFiles(WixSetupOptions options)
        {
            var docWxs = new XmlDocument();
            var nodeWxs = docWxs.CreateElement("root");
            docWxs.AppendChild(nodeWxs);

            var docWixProj = new XmlDocument();
            var nodeWixProj = docWixProj.CreateElement("root");
            docWixProj.AppendChild(nodeWixProj);

            const string PublishTemplateItemName = "Publish6EACFB1ABA5A4581A2F0DFA55A8B3445";
            const string PublishStencilItemName = "PublishE8358BB3898744BEA3D6E8B0DE0D80F4";

            var wxs = "";
            var wixProj = "";

            if (options.EnableLicense)
            {
                var nodeLicense = docWixProj.CreateElement("Content");
                nodeLicense.SetAttribute("Include", Path.GetFileName(options.LicenseFilePath));
                nodeWixProj.AppendChild(nodeLicense);
            }

            if (options.HaveVisioFiles)
            {
                foreach (var path in options.VisioFilePaths)
                {
                    var nodeComponent = docWxs.CreateElement("Component");
                    nodeWxs.AppendChild(nodeComponent);

                    var nodeContent = docWixProj.CreateElement("Content");
                    nodeContent.SetAttribute("Include", Path.GetFileName(path));
                    nodeWixProj.AppendChild(nodeContent);

                    var nodeFile = docWxs.CreateElement("File");
                    nodeFile.SetAttribute("Name", Path.GetFileName(path));

                    if (!options.DuplicateExistingVisioFiles)
                        nodeFile.SetAttribute("Source", path);

                    nodeComponent.AppendChild(nodeFile);

                    string ext = Path.GetExtension(path);

                    string elementName = null;

                    if (ext == ".vss" || ext == ".vssx" || ext == ".vssm" || ext == ".vsx")
                        elementName = PublishStencilItemName;

                    if (ext == ".vst" || ext == ".vstx" || ext == ".vstm" || ext == ".vtx")
                        elementName = PublishTemplateItemName;

                    if (elementName != null)
                    {
                        var nodePublish = docWxs.CreateElement(elementName);

                        nodePublish.SetAttribute("MenuPath", string.Format("{0}\\{1}",
                            GlobalDictionary["$csprojectname$"], Path.GetFileNameWithoutExtension(path)));

                        nodeFile.AppendChild(nodePublish);
                    }
                }
            }

            wxs = beautifyXml(docWxs)
                .Replace(PublishTemplateItemName, "visio:PublishTemplate")
                .Replace(PublishStencilItemName, "visio:PublishStencil")
                .Replace("<root>", "")
                .Replace("<root />", "")
                .Replace("</root>", "");

            wixProj = beautifyXml(docWixProj)
                .Replace("<root>", "")
                .Replace("<root />", "")
                .Replace("</root>", "");

            GlobalDictionary["$visioFilesWxs$"] = wxs;
            GlobalDictionary["$visioFilesWixProj$"] = wixProj;
        }

        public bool IsWixInstalled()
        {
            IServiceProvider serviceProvider = new ServiceProvider(_dte as Microsoft.VisualStudio.OLE.Interop.IServiceProvider);

            var shell = (IVsShell)serviceProvider.GetService(typeof(SVsShell));
            if (shell == null)
                return true;

            int wixInstalled;
            var wixGuid = new Guid("E0EE8E7D-F498-459E-9E90-2B3D73124AD5");
            if (0 != shell.IsPackageInstalled(ref wixGuid, out wixInstalled))
                return true;

            return wixInstalled != 0;
        }

        static void GetVisioPath(RegistryKey key, string version, ref string path)
        {
            var subKey = key.OpenSubKey(string.Format(@"Software\Microsoft\Office\{0}\Visio\InstallRoot", version));
            if (subKey == null)
                return;

            var value = subKey.GetValue("Path", null);
            if (value == null)
                return;

            path = Path.Combine(value.ToString(), "Visio.exe");
        }

        public static string GetVisioPath32()
        {
            var key = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry32);
            string path = null;
            foreach (var item in new[] { "11.0", "12.0", "14.0", "15.0", "16.0" })
                GetVisioPath(key, item, ref path);
            return path;
        }

        public static string GetVisioPath64()
        {
            if (!Environment.Is64BitOperatingSystem)
                return null;

            var key = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64);
            string path = null;
            foreach (var item in new[] { "14.0", "15.0", "16.0" })
                GetVisioPath(key, item, ref path);
            return path;
        }

        // Don't return the latest version, even if it is installed, because it will cause problems by auto-upgrade
        static string GetOfficeVersion()
        {
            return Registry.GetValue(@"HKEY_LOCAL_MACHINE\Software\Microsoft\Office\14.0\Visio\InstallRoot", "Path", null) != null
                ? "14.0" : "12.0";
        }

        void SetActiveConfiguration()
        {
            try
            {
                var x64 = GetVisioPath64() != null;

                foreach (SolutionConfiguration2 config in _dte.Solution.SolutionBuild.SolutionConfigurations)
                {
                    if (config.PlatformName == "x64" || config.PlatformName == "x86")
                    {
                        foreach (SolutionContext context in config.SolutionContexts)
                            context.ShouldBuild = true;
                    }

                    if (config.Name != "Debug")
                        continue;

                    if (x64 && config.PlatformName == "x64")
                    {
                        config.Activate();
                        break;
                    }

                    if (!x64 && config.PlatformName == "x86")
                    {
                        config.Activate();
                        break;
                    }
                }
            }
            // this convenience feature; continue if failed, not a big deal
            // ReSharper disable once EmptyGeneralCatchClause
            catch (Exception)
            {
            }
        }

        public void RunFinished()
        {
            SetActiveConfiguration();
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

        public bool IsVstoInstalled()
        {
            var vstoKey = string.Format(@"HKEY_LOCAL_MACHINE\{0}\Setup\VSTO", _dte.RegistryRoot);
            return Registry.GetValue(vstoKey, "ProductDir", null) != null;
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
                    return "http://wixtoolset.org/releases/";

                case ExternalLink.VstoDownload:
                {
                    var match = Regex.Match(_dte.Version, @"\d+");
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
}
