Imports Microsoft.Office.Tools.Ribbon

Public Class AddinRibbonComponent

    Private Sub Command1_Click(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles Command1.Click
        Globals.ThisAddIn.Command1()
    End Sub
    $if$ ($taskpane$ == true)
    Private Sub TogglePanel_Click(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles TogglePanel.Click
        Globals.ThisAddIn.TogglePanel()
    End Sub
$endif$
End Class
