Public Class ThisAddIn

    $if$ ($taskpaneORui$ == true)Private ReadOnly _addin As Addin = New Addin()
    $endif$
    $if$ ($ribbon$ == true)
    Protected Overrides Function CreateRibbonExtensibilityObject() As Office.IRibbonExtensibility
        Return _addin
    End Function
    $endif$
    Private Sub ThisAddIn_Startup() Handles Me.Startup
        $if$ ($taskpaneORui$ == true)_addin.Startup(Application)
    $endif$End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown
        $if$ ($taskpaneORui$ == true)_addin.Shutdown()
    $endif$End Sub

End Class
