namespace PanelAddinWizard
{
    public enum ExternalLink
    {
        WixDownload,
        WixDocsUI,
        VstoDownload,
    }

    public interface IWizardFormHost
    {
        bool IsWixInstalled();
        bool IsVstoInstalled();
        void OpenExternalLink(ExternalLink link);
    }
}
