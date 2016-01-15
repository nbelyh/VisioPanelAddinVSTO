using System;
using System.IO.MemoryMappedFiles;
using System.Threading;
using System.Xml.Serialization;

namespace PanelAddinWizard
{
    public class XmlWizardOptionsManager
    {
        public static void PanelAddinWizardTestApp(Action action)
        {
            using (new Mutex(false, @"Global\PanelAddinWizardTestApp"))
            {
                action();
            }
        }

        public static bool IsPanelAddinWizardTestAppStarted()
        {
            bool created;
            using (new Mutex(false, @"Global\PanelAddinWizardTestApp", out created))
                return !created;
        }

        public static XmlWizardOptions Read()
        {
            using (var f = MemoryMappedFile.OpenExisting(@"Global\PanelAddinWizardTestAppFile"))
            using (var fs = f.CreateViewStream())
            {
                var serializer = new XmlSerializer(typeof (XmlWizardOptions));
                return (XmlWizardOptions) serializer.Deserialize(fs);
            }
        }

        public static void Write(XmlWizardOptions options, Action action)
        {
            using (var f = MemoryMappedFile.CreateNew(@"Global\PanelAddinWizardTestAppFile", 4096))
            using (var fs = f.CreateViewStream())
            {
                var serializer = new XmlSerializer(typeof(XmlWizardOptions));
                serializer.Serialize(fs, options);

                action();
            }
        }
    }

    [Serializable]
    public class XmlWizardOptions : IWizardOptions
    {
        public bool AddinEnabled { get; set; }
        public bool InstallExtensibilityInterop { get; set; }
        public bool InstallVisioInterops { get; set; }
        public bool TaskPane { get; set; }
        public bool CommandBars { get; set; }
        public bool RibbonXml { get; set; }
        public bool RibbonComponent { get; set; }
        public bool AddinTypeCOM { get; set; }
        public bool AddinTypeVSTO { get; set; }
        public bool EnableWixSetup { get; set; }
        public bool AddVisioFiles { get; set; }
        public bool CreateNewVisioFiles { get; set; }
        public bool DuplicateExistingVisioFiles { get; set; }
        public bool EnableWixUI { get; set; }

        public string AddinName { get; set; }
        public string AddinDescription { get; set; }
        public string WixUI { get; set; }

        public string[] VisioFilePaths { get; set; }
    }
}