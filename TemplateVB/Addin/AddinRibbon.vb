Imports System.Drawing
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Core


''' 
''' User interface manager for Visio 2010 and above
''' Creates and controls ribbon UI
''' 
$if$ ($vstoAddin$ == true)<ComVisible(True)>
$endif$Public Partial Class $if$ ($vstoAddin$ == true)AddinUI$else$ThisAddIn$endif$
    Implements IRibbonExtensibility
	Private _ribbon As Microsoft.Office.Core.IRibbonUI

#Region "IRibbonExtensibility Members"

    Public Function GetCustomUI(ribbonId As String) As String Implements IRibbonExtensibility.GetCustomUI
        Return My.Resources.Ribbon
    End Function

#End Region

#Region "Ribbon Callbacks"
    
    Public Function IsRibbonCommandEnabled(ctrl As Microsoft.Office.Core.IRibbonControl) As Boolean
            Return $thisAddIn$IsCommandEnabled(ctrl.Id)
    End Function
    
    Public Function IsRibbonCommandChecked(ctrl As Microsoft.Office.Core.IRibbonControl) As Boolean
            Return $thisAddIn$IsCommandChecked(ctrl.Id)
    End Function
    
    Public Sub OnRibbonButtonCheckClick(control As Microsoft.Office.Core.IRibbonControl, pressed As Boolean)
            $thisAddIn$OnCommand(control.Id)
    End Sub
    
    Public Sub OnRibbonButtonClick(control As Microsoft.Office.Core.IRibbonControl)
            $thisAddIn$OnCommand(control.Id)
    End Sub
    
    Public Function OnGetRibbonLabel(control As Microsoft.Office.Core.IRibbonControl) As String
            Return $thisAddIn$GetCommandLabel(control.Id)
    End Function
    
    Public Sub OnRibbonLoad(ribbonUI As Microsoft.Office.Core.IRibbonUI)
        _ribbon = ribbonUI
    End Sub
    
    Public Function GetRibbonImage(control As Microsoft.Office.Core.IRibbonControl) As Bitmap
            Return $thisAddIn$GetCommandBitmap(control.Id)
    End Function
    
    #End Region
    
    Public Sub UpdateRibbon()
        _ribbon.Invalidate()
    End Sub

End Class