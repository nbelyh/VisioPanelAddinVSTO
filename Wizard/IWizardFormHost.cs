namespace PanelAddinWizard
{
    public interface IWizardFormHost
    {
        bool IsWixInstalled();
        bool IsVstoInstalled();
        int GetVisualStudioVersion();
    }
}
