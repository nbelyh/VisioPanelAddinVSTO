$if$ ($uiCallbacks$ == true)Imports System.Drawing
$endif$$if$ ($ribbonANDcommandbars$ == true)Imports System.Globalization
$endif$$if$ ($ui$ == true)Imports System.Windows.Forms
$endif$$if$ ($ui$ == true)Imports Visio = Microsoft.Office.Interop.Visio
$endif$$if$ ($uiCallbacks$ == true)Imports System.Runtime.InteropServices
$endif$
Public Class ThisAddIn

    $if$ ($uiCallbacks$ == true)Private ReadOnly AddinUI As AddinUI = New AddinUI()
    $endif$$if$ ($ribbonXml$ == true)
    Protected Overrides Function CreateRibbonExtensibilityObject() As Office.IRibbonExtensibility
        Return _addinUI
    End Function
    $endif$$if$ ($ui$ == true)
    ''' 
    ''' A simple command
    ''' 
    Public Sub Command1()
        MessageBox.Show("Hello from command 1!")
    End Sub
    $endif$$if$ ($uiCallbacks$ == true)
    ''' 
    ''' A command to demonstrate conditionally enabling/disabling.
    ''' The command gets enabled only when a shape is selected
    ''' 
    Public Sub Command2()
        If Application Is Nothing OrElse Application.ActiveWindow Is Nothing OrElse Application.ActiveWindow.Selection Is Nothing Then Exit Sub

        MessageBox.Show(String.Format("Hello from (conditional) command 2! You have {0} shapes selected.", Application.ActiveWindow.Selection.Count))
    End Sub
    $endif$$if$ ($uiCallbacks$ == true)
    ''' 
    ''' Callback called by the UI manager when user clicks a button
    ''' Should do something meaninful wehn corresponding action is called.
    ''' 
    Public Sub OnCommand(commandId As String)
        Select Case commandId
            Case "Command1"
                Command1()
                Return
            $endif$$if$ ($uiCallbacks$ == true)
            Case "Command2"
                Command2()
                Return
			$endif$$if$ ($taskpaneANDuiCallbacks$ == true)
            Case "TogglePanel"
                TogglePanel()
                Return
        $endif$$if$ ($uiCallbacks$ == true)
        End Select
    End Sub
    $endif$$if$ ($uiCallbacks$ == true)
    ''' 
    ''' Callback called by UI manager.
    ''' Should return if corresponding command shoudl be enabled in the user interface.
    ''' By default, all commands are enabled.
    ''' 
    Public Function IsCommandEnabled(commandId As String) As Boolean
        Select Case commandId
            Case "Command1"
                ' make command1 always enabled
                Return True

            Case "Command2"
                ' make command2 enabled only if a window is opened
                Return Application IsNot Nothing AndAlso Application.ActiveWindow IsNot Nothing AndAlso Application.ActiveWindow.Selection.Count > 0
			$endif$$if$ ($taskpaneANDuiCallbacks$ == true)
            Case "TogglePanel"
                ' make panel enabled only if we have selected shape(s).
                Return IsPanelEnabled()
            $endif$$if$ ($uiCallbacks$ == true)Case Else
                Return True
        End Select
    End Function

    ''' 
    ''' Callback called by UI manager.
    ''' Should return if corresponding command (button) is pressed or not (makes sense for toggle buttons)
    ''' 
    Public Function IsCommandChecked(command As String) As Boolean
		$endif$$if$ ($taskpaneANDuiCallbacks$ == true)
        If command = "TogglePanel" Then
            Return IsPanelVisible()
        End If

        $endif$$if$($uiCallbacks$ == true)Return False
    End Function

    ''' 
    ''' Callback called by UI manager.
    ''' Returns a label associated with given command.
    ''' We assume for simplicity taht command labels are named simply named as [commandId]_Label (see resources)
    ''' 
    Public Function GetCommandLabel(command As String) As String
        Return My.Resources.ResourceManager.GetString(command & "_Label")
    End Function

    ''' 
    ''' Returns a icon associated with given command.
    ''' We assume for simplicity that icon ids are named after command commandId.
    ''' 
    Public Function GetCommandBitmap(command As String) As Bitmap
        Return DirectCast(My.Resources.ResourceManager.GetObject(command), Bitmap)
    End Function
	$endif$$if$ ($taskpane$ == true)
    Public Sub TogglePanel()
        _panelManager.TogglePanel(Application.ActiveWindow)
    End Sub
    $endif$$if$ ($taskpaneANDuiCallbacks$ == true)
    Public Function IsPanelEnabled() As Boolean
        Return Application IsNot Nothing AndAlso Application.ActiveWindow IsNot Nothing
    End Function

    Public Function IsPanelVisible() As Boolean
        Return Application IsNot Nothing AndAlso _panelManager.IsPanelOpened(Application.ActiveWindow)
    End Function
    $endif$$if$ ($taskpane$ == true)
    Private _panelManager As PanelManager
	$endif$
	$if$ ($uiCallbacks$ == true)
    Sub UpdateUI()
        $endif$$if$ ($commandbars$ == true)AddinUI.UpdateCommandBars()
        $endif$$if$ ($ribbonXml$ == true)AddinUI.UpdateRibbon()
    $endif$$if$ ($uiCallbacks$ == true)
    End Sub
    $endif$$if$ ($uiCallbacks$ == true)
    Public Sub Application_SelectionChanged(window As Visio.Window)
        UpdateUI()
    End Sub
    $endif$
    Private Sub ThisAddIn_Startup() Handles Me.Startup
        $if$ ($taskpane$ == true)_panelManager = New PanelManager(Me)
		$endif$$if$ ($ribbonANDcommandbars$ == true)Dim version = Integer.Parse(Application.Version, NumberStyles.AllowDecimalPoint)
        If (version < 14) Then
			$endif$$if$ ($commandbars$ == true)AddinUI.StartupCommandBars("$csprojectname$", New String() {"Command1", "Command2"$endif$$if$ ($commandbarsANDtaskpane$ == true), "TogglePanel"$endif$$if$ ($commandbars$ == true)})
        $endif$$if$ ($ribbonANDcommandbars$ == true)End If
        $endif$$if$ ($uiCallbacks$ == true)AddHandler Application.SelectionChanged, AddressOf Application_SelectionChanged
        $endif$
    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown
        $if$ ($commandbars$ == true)AddinUI.ShutdownCommandBars()
        $endif$$if$ ($taskpane$ == true)_panelManager.Dispose()
        $endif$$if$ ($uiCallbacks$ == true)RemoveHandler Application.SelectionChanged, AddressOf Application_SelectionChanged
        $endif$
    End Sub

End Class
$if$ ($uiCallbacks$ == true)
<ComVisible(True)>
Partial Public Class AddinUI
    ReadOnly Property ThisAddIn() As ThisAddIn
        Get
            Return Globals.ThisAddIn
        End Get
    End Property
End Class
$endif$