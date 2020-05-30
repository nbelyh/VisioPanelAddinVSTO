using System;
using System.IO;
using EnvDTE;
using EnvDTE80;
using PanelAddinWizard;
using System.Xml.Serialization;
using System.Xml;

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

        static void TestAddin(XmlWizardOptions options)
        {
            options.AddinTypeCOM = true;
            options.AddinTypeVSTO = false;
            TestAddinCOM(options);

            TestInterop(options);

            options.AddinTypeCOM = false;
            options.AddinTypeVSTO = true;
            TestAddinVSTO(options);

            TestInterop(options);
        }

        static void TestInterop(XmlWizardOptions options)
        {
            options.InstallExtensibilityInterop = false;
            options.InstallVisioInterops = true;
            ExecuteInstall(options);

            options.InstallExtensibilityInterop = true;
            options.InstallVisioInterops = false;
            ExecuteInstall(options);

            options.InstallExtensibilityInterop = true;
            options.InstallVisioInterops = true;
            ExecuteInstall(options);

            options.InstallExtensibilityInterop = false;
            options.InstallVisioInterops = false;
        }

        static void TestAddinCOM(XmlWizardOptions options)
        {
            options.RibbonXml = false;
            options.RibbonComponent = false;
            TaskPane(options);

            options.RibbonXml = true;
            options.RibbonComponent = false;
            TaskPane(options);

            options.RibbonXml = false;
            options.RibbonComponent = false;
        }

        static void TestAddinVSTO(XmlWizardOptions options)
        {
            TestAddinCOM(options);

            options.RibbonXml = true;
            options.RibbonComponent = true;
            TaskPane(options);

            options.RibbonXml = false;
            options.RibbonComponent = false;
        }

        static void TaskPane(XmlWizardOptions options)
        {
            options.TaskPane = false;
            CommandBars(options);

            options.TaskPane = true;
            CommandBars(options);
        }

        static void CommandBars(XmlWizardOptions options)
        {
            options.CommandBars = false;
            ExecuteInstall(options);

            options.CommandBars = true;
            ExecuteInstall(options);
        }

        static void TestSetup(XmlWizardOptions options)
        {
            options.AddinTypeCOM = true;
            options.AddinTypeVSTO = false;
            TestSetupType(options);

            options.AddinTypeCOM = false;
            options.AddinTypeVSTO = true;
            TestSetupType(options);
        }

        static void TestSetupType(XmlWizardOptions options)
        {
            TestUI(options);
            TestVisioFiles(options);
        }

        static void TestUI(XmlWizardOptions options)
        {
            options.EnableWixUI = false;
            ExecuteInstall(options);

            options.EnableWixUI = true;

            options.WixUI = "WixUI_Minimal";
            ExecuteInstall(options);

            options.WixUI = "WixUI_InstallDir";
            ExecuteInstall(options);

            options.WixUI = "WixUI_InstallDirNoLicense";
            ExecuteInstall(options);

            options.WixUI = "WixUI_Mondo";
            ExecuteInstall(options);

            options.WixUI = "WixUI_Advanced";
            ExecuteInstall(options);
        }

        static void TestVisioFiles(XmlWizardOptions options)
        {
            options.AddVisioFiles = true;
            options.DuplicateExistingVisioFiles = false;

            options.CreateNewVisioFiles = true;
            ExecuteInstall(options);

            options.CreateNewVisioFiles = false;
            options.VisioFilePaths = new[]
            {
                @"C:\Projects\github\VisioPanelAddinVSTO\PanelAddinWizardTestApp\Data\X1_M.vss",
                @"C:\Projects\github\VisioPanelAddinVSTO\PanelAddinWizardTestApp\Data\X1_M.vst",
                @"C:\Projects\github\VisioPanelAddinVSTO\PanelAddinWizardTestApp\Data\X2_M.vss",
                @"C:\Projects\github\VisioPanelAddinVSTO\PanelAddinWizardTestApp\Data\X2_M.vst"
            };
            ExecuteInstall(options);

            TestLanguage(options);

            options.DuplicateExistingVisioFiles = true;
            ExecuteInstall(options);
        }

        static void TestLanguage(XmlWizardOptions options)
        {
            options.UseSetupLanguage = true;
            options.SetupLanguage = "ru-RU";
            options.VisioFilePaths = new[]
            {
                @"C:\Projects\github\VisioPanelAddinVSTO\PanelAddinWizardTestApp\Data\фыва.vss",
                @"C:\Projects\github\VisioPanelAddinVSTO\PanelAddinWizardTestApp\Data\фыва.vst",
                @"C:\Projects\github\VisioPanelAddinVSTO\PanelAddinWizardTestApp\Data\йцук.vss",
                @"C:\Projects\github\VisioPanelAddinVSTO\PanelAddinWizardTestApp\Data\йцук.vst"
            };
            ExecuteInstall(options);

            options.UseSetupLanguage = false;
            options.SetupLanguage = "";
            options.VisioFilePaths = null;
        }

        static void ExecuteInstall(XmlWizardOptions options)
        {
            XmlWizardOptionsManager.Write(options, () =>
            {
                var name = "addin_" + Guid.NewGuid().ToString("N");

                var sln = (Solution2)DTE.Solution;

                var templatePath = sln.GetProjectTemplate("Template.zip", "VisualBasic");

                var path = Path.Combine(TestPath, name);
                sln.AddFromTemplate(templatePath, path, name);

                var serializer = new XmlSerializer(typeof(XmlWizardOptions));
                using (var streamWriter = new StreamWriter(Path.Combine(path, "settings.xml")))
                using (var xmlWriter = XmlWriter.Create(streamWriter, new XmlWriterSettings { Indent = true}))
                {
                    serializer.Serialize(xmlWriter, options);
                }

                foreach (SolutionConfiguration2 sc2 in sln.SolutionBuild.SolutionConfigurations)
                    foreach (SolutionContext sc in sc2.SolutionContexts)
                        sc.ShouldBuild = true;

                DTE.ExecuteCommand("Build.BuildSolution");

                while (DTE.Solution.SolutionBuild.BuildState != vsBuildState.vsBuildStateDone)
                    System.Threading.Thread.Sleep(500);

                var failed = sln.SolutionBuild.LastBuildInfo > 0;
                sln.Close();

                if (!failed)
                {
                    for (int i = 0; i < 10; ++i)
                    {
                        try
                        {
                            Directory.Delete(path, true);
                            break;
                        }
                        catch
                        {
                            System.Threading.Thread.Sleep(1000);
                        }
                    }
                }

            });
        }

        private static DTE2 DTE;
        private static string TestPath;

        private static void DoTests(string vs)
        {
            var type = Type.GetTypeFromProgID(vs);
            DTE = (DTE2)Activator.CreateInstance(type);

            TestPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, vs);
            if (!Directory.Exists(TestPath))
                Directory.CreateDirectory(TestPath);

            MessageFilter.Register();

            DTE.MainWindow.Activate();

            var options = new XmlWizardOptions();
            options.AddinName = "_Name";
            options.AddinManufacturer = "_AddinManufacturer";
            options.AddinDescription = "_Description";

            options.AddinEnabled = true;
            TestAddin(options);

            options.EnableWixSetup = true;
            TestSetup(options);

            options.AddinTypeCOM = false;
            options.AddinTypeVSTO = false;

            options.EnableWixSetup = true;
            options.AddinEnabled = false;
            TestVisioFiles(options);

            DTE.Quit();
            MessageFilter.Revoke();
        }

        [STAThread]
        static void Main(string[] args)
        {
            var wizardForm = new WizardForm(new TestHost(), "TestAddIn");
            wizardForm.ShowDialog();

            //XmlWizardOptionsManager.PanelAddinWizardTestApp(() => DoTests("VisualStudio.DTE.10.0"));

            //XmlWizardOptionsManager.PanelAddinWizardTestApp(() => DoTests("VisualStudio.DTE.12.0"));

            //XmlWizardOptionsManager.PanelAddinWizardTestApp(() => DoTests("VisualStudio.DTE.14.0"));
        }
    }
}
