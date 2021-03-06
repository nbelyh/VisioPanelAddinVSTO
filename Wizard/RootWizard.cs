﻿
using System;
using System.CodeDom;
using System.CodeDom.Compiler;
using System.IO;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Drawing;
using System.Security;
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

        public static IWizardOptions WizardOptions;

        private DTE2 _dte;

        protected abstract Image HeaderImage { get; }

        IWizardOptions ConfigureOptionsSource(string projectName)
        {
#if DEBUG
            if (XmlWizardOptionsManager.IsPanelAddinWizardTestAppStarted())
                return XmlWizardOptionsManager.Read();
#endif
            var form = new WizardForm(this, projectName)
            {
                HeaderImage = HeaderImage
            };

            if (form.ShowDialog() == DialogResult.Cancel)
                throw new WizardBackoutException();

            return form;
        }

        // Add global replacement parameters
        public void RunStarted(object automationObject,
            Dictionary<string, string> replacementsDictionary,
            WizardRunKind runKind, object[] customParams)
        {
            _dte = automationObject as DTE2;

            WizardOptions = ConfigureOptionsSource(replacementsDictionary["$safeprojectname$"]);

            GlobalDictionary["$csprojectname$"] = replacementsDictionary["$safeprojectname$"];

            var addinManufacturer = WizardOptions.AddinManufacturer;
            if (string.IsNullOrEmpty(addinManufacturer))
                addinManufacturer = Environment.UserName;

            GlobalDictionary["$addinManufacturerXML$"] = EncodeXML(addinManufacturer);

            var productName = WizardOptions.AddinName;
            if (string.IsNullOrEmpty(productName))
                productName = replacementsDictionary["$safeprojectname$"];

            GlobalDictionary["$productNameXML$"] = EncodeXML(productName);
            GlobalDictionary["$productNameCS$"] = EncodeForCS(productName);
            GlobalDictionary["$productNameVB$"] = EncodeForVB(productName);

            GlobalDictionary["$progid$"] = replacementsDictionary["$safeprojectname$"] + ".Addin";
            GlobalDictionary["$clsid$"] = replacementsDictionary["$guid1$"];
            GlobalDictionary["$wixproject$"] = replacementsDictionary["$guid2$"];
            GlobalDictionary["$csprojectguid$"] = replacementsDictionary["$guid3$"];

            GlobalDictionary["$mergeguid$"] = replacementsDictionary["$guid4$"];

            var ribbon = WizardOptions.RibbonXml || WizardOptions.RibbonComponent;

            GlobalDictionary["$ribbon$"] = ribbon ? "true" : "false";
            GlobalDictionary["$ribbonXml$"] = WizardOptions.RibbonXml ? "true" : "false";
            GlobalDictionary["$ribbonComponent$"] = WizardOptions.RibbonComponent ? "true" : "false";

            GlobalDictionary["$commandbars$"] = WizardOptions.CommandBars ? "true" : "false";

            GlobalDictionary["$ribbonANDcommandbars$"] = ribbon  && WizardOptions.CommandBars ? "true" : "false";
            GlobalDictionary["$commandbarsANDtaskpane$"] = WizardOptions.CommandBars && WizardOptions.TaskPane ? "true" : "false";
            GlobalDictionary["$taskpane$"] = WizardOptions.TaskPane ? "true" : "false";

            var ui = (WizardOptions.CommandBars || ribbon);
            GlobalDictionary["$ui$"] = ui ? "true" : "false";
            GlobalDictionary["$taskpaneANDui$"] = (WizardOptions.TaskPane && ui) ? "true" : "false";
            GlobalDictionary["$taskpaneORui$"] = (WizardOptions.TaskPane || ui) ? "true" : "false";

            var uiCallbacks = (WizardOptions.CommandBars || WizardOptions.RibbonXml);
            GlobalDictionary["$uiCallbacks$"] = uiCallbacks ? "true" : "false";
            GlobalDictionary["$taskpaneANDuiCallbacks$"] = (WizardOptions.TaskPane && uiCallbacks) ? "true" : "false";

            GlobalDictionary["$addinProject$"] = WizardOptions.AddinEnabled ? "true" : "false";
            GlobalDictionary["$wixSetup$"] = WizardOptions.EnableWixSetup ? "true" : "false";

            GetFiles();

            GlobalDictionary["$EnableWixUI$"] = WizardOptions.EnableWixUI ? "true" : "false";
            GlobalDictionary["$WixUI$"] = WizardOptions.EnableWixUI ? WizardOptions.WixUI : "";
            GlobalDictionary["$addVisioFiles$"] = WizardOptions.EnableWixSetup && WizardOptions.AddVisioFiles ? "true" : "false";
            GlobalDictionary["$defaultVisioFiles$"] = WizardOptions.EnableWixSetup && WizardOptions.CreateNewVisioFiles ? "true" : "false";

            GlobalDictionary["$vstoAddin$"] = WizardOptions.AddinTypeVSTO ? "true" : "false";
            GlobalDictionary["$comAddin$"] = WizardOptions.AddinTypeCOM ? "true" : "false";
            GlobalDictionary["$registerForComInterop$"] = WizardOptions.AddinTypeCOM ? "true" : "false";

            GlobalDictionary["$ribbonXmlVSTO$"] = WizardOptions.RibbonXml && WizardOptions.AddinTypeVSTO ? "true" : "false";
            GlobalDictionary["$thisAddIn$"] = WizardOptions.AddinTypeVSTO ? "Globals.ThisAddIn." : "";
            GlobalDictionary["$thisAddInUI$"] = WizardOptions.AddinTypeVSTO ? "AddinUI." : "";
            GlobalDictionary["$uiCallbacksVSTO$"] = WizardOptions.AddinTypeVSTO && uiCallbacks ? "true" : "";

            GlobalDictionary["$installExtensibilityInterop$"] = WizardOptions.InstallExtensibilityInterop ? "true" : "false";
            GlobalDictionary["$installVisioInterops$"] = WizardOptions.InstallVisioInterops ? "true" : "false";

            GlobalDictionary["$AddinFriendlyNameVB$"] = EncodeForVB(WizardOptions.AddinName);
            GlobalDictionary["$AddinFriendlyNameCS$"] = EncodeForCS(WizardOptions.AddinName);

            GlobalDictionary["$AddinManufactorerVB$"] = EncodeForVB(WizardOptions.AddinManufacturer);
            GlobalDictionary["$AddinManufactorerCS$"] = EncodeForCS(WizardOptions.AddinManufacturer);

            GlobalDictionary["$AddinDescriptionVB$"] = EncodeForVB(WizardOptions.AddinDescription);
            GlobalDictionary["$AddinDescriptionCS$"] = EncodeForCS(WizardOptions.AddinDescription);

            GlobalDictionary["$vstoHostPackageGuid$"] = GetVstoHostPackageGuid();
            GlobalDictionary["$vstoTargetOfficeVersion$"] = GetTargetVstoOfficeVersion();
            GlobalDictionary["$vstoTargetExeVersion$"] = GetTargetVstoExeVersion();

            GlobalDictionary["$SetupLanguage$"] = WizardOptions.UseSetupLanguage ? WizardOptions.SetupLanguage : "";

            GlobalDictionary["$vsToolsVersion$"] = GetVsToolsVersion();
        }

        string GetVstoHostPackageGuid()
        {
            return GetVisualStudioVersion() < 12
                ? "{20A848B8-E01F-4801-962E-25DB0FF57389}"
                : "{29A7B9D7-A7F1-4328-8EF0-6B2D1A56B2C1}";
        }

        string GetTargetVstoOfficeVersion()
        {
            return GetVisualStudioVersion() < 12
                ? "14.0"
                : "15.0";
        }

        string GetTargetVstoExeVersion()
        {
            foreach (var item in new[] { "16.0", "15.0", "14.0", "12.0", "11.0" })
            {
                var keyPath = string.Format(@"HKEY_LOCAL_MACHINE\Software\Microsoft\Office\{0}\Visio\InstallRoot", item);
                if (Registry.GetValue(keyPath, "Path", null) != null)
                    return item;
            }

            return GetTargetVstoOfficeVersion();
        }

        string GetVsToolsVersion()
        {
            switch (GetVisualStudioVersion())
            {
                case 12:
                    return "12.0";
                case 14:
                    return "14.0";
                default:
                    return "4.0";
            }
        }

        private static string EncodeForCS(string input)
        {
            return EncodeForCode(input, "CSharp");
        }

        private static string EncodeForVB(string input)
        {
            return EncodeForCode(input, "VisualBasic");
        }

        private static string EncodeForCode(string input, string providerName)
        {
            using (var writer = new StringWriter())
            {
                using (var provider = CodeDomProvider.CreateProvider(providerName))
                {
                    provider.GenerateCodeFromExpression(new CodePrimitiveExpression(input), writer, null);
                    return writer.ToString();
                }
            }
        }

        private static string EncodeXML(string productName)
        {
            return SecurityElement.Escape(productName);
        }

        private static string BeautifyXml(XmlDocument doc)
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

        public static bool HaveVisioFiles 
        {
            get
            {
                return WizardOptions.EnableWixSetup && !WizardOptions.CreateNewVisioFiles && WizardOptions.VisioFilePaths != null && WizardOptions.VisioFilePaths.Length > 0;
            } 
        }

        void GetFiles()
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

            if (HaveVisioFiles)
            {
                foreach (var path in WizardOptions.VisioFilePaths)
                {
                    var nodeComponent = docWxs.CreateElement("Component");
                    nodeWxs.AppendChild(nodeComponent);

                    var nodeContent = docWixProj.CreateElement("Content");
                    nodeContent.SetAttribute("Include", Path.GetFileName(path));
                    nodeWixProj.AppendChild(nodeContent);

                    var nodeFile = docWxs.CreateElement("File");

                    var fileName = Path.GetFileName(path);

                    nodeFile.SetAttribute("Name", fileName);

                    // for non-english, take care about Id
                    if (WizardOptions.UseSetupLanguage)
                    {
                        var id = Translit.MakeIdentifierFromString(fileName);
                        if (id != fileName)
                            nodeFile.SetAttribute("Id", id);
                    }

                    if (!WizardOptions.DuplicateExistingVisioFiles)
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
                            WizardOptions.AddinName, Path.GetFileNameWithoutExtension(path)));

                        nodeFile.AppendChild(nodePublish);
                    }
                }
            }

            wxs = BeautifyXml(docWxs)
                .Replace(PublishTemplateItemName, "visio:PublishTemplate")
                .Replace(PublishStencilItemName, "visio:PublishStencil")
                .Replace("<root>", "")
                .Replace("<root />", "")
                .Replace("</root>", "");

            wixProj = BeautifyXml(docWixProj)
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

        public int GetVisualStudioVersion()
        {
            var match = Regex.Match(_dte.Version, @"\d+");
            return match.Success ? int.Parse(match.Value) : 0;
        }
    }
}
