namespace PanelAddinWizard
{
    public interface IWizardOptions
    {
        bool AddinEnabled { get; }

        bool InstallExtensibilityInterop { get; }
        bool InstallVisioInterops { get; }
        bool TaskPane { get; }
        bool CommandBars { get; }

        bool RibbonXml { get; }
        bool RibbonComponent { get; }

        bool AddinTypeCOM { get; }
        bool AddinTypeVSTO { get; }

        string AddinName { get; }
        string AddinDescription { get; }

        bool EnableWixSetup { get; }
        bool AddVisioFiles { get; }
        bool CreateNewVisioFiles { get; }
        bool DuplicateExistingVisioFiles { get; }
        string[] VisioFilePaths { get; }
        bool EnableWixUI { get; }
        string WixUI { get; }
    }
}
