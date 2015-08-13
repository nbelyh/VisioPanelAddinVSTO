﻿
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

namespace PanelAddinWizard
{
    /// <summary>
    /// Summary description for WizardForm.
    /// </summary>
    public abstract class RootWizard
        : IWizard
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

            var wixInstalled = IsWixInstalled();

            var wizardForm = new WizardForm(wixInstalled)
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

            GlobalDictionary["$ribbon$"] = wizardForm.Ribbon ? "true" : "false";
            GlobalDictionary["$commandbars$"] = wizardForm.CommandBars ? "true" : "false";

            GlobalDictionary["$ribbonORcommandbars$"] = wizardForm.Ribbon || wizardForm.CommandBars ? "true" : "false";
            GlobalDictionary["$ribbonANDcommandbars$"] = wizardForm.Ribbon && wizardForm.CommandBars ? "true" : "false";
            GlobalDictionary["$commandbarsANDtaskpane$"] = wizardForm.CommandBars && wizardForm.TaskPane ? "true" : "false";
            GlobalDictionary["$taskpane$"] = wizardForm.TaskPane ? "true" : "false";
            GlobalDictionary["$ui$"] = (wizardForm.CommandBars || wizardForm.Ribbon) ? "true" : "false";
            GlobalDictionary["$taskpaneANDui$"] = (wizardForm.TaskPane && (wizardForm.CommandBars || wizardForm.Ribbon)) ? "true" : "false";
            GlobalDictionary["$taskpaneORui$"] = (wizardForm.TaskPane || (wizardForm.CommandBars || wizardForm.Ribbon)) ? "true" : "false";

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

        private bool IsWixInstalled()
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
