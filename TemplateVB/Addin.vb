$if$ ($ui$ == true)Imports System.Drawing
$endif$$if$ ($ribbonANDcommandbars$ == true)Imports System.Globalization
$endif$$if$ ($ui$ == true)Imports System.Windows.Forms
$endif$

Partial Public Class Addin
    Public Property Application As Microsoft.Office.Interop.Visio.Application
    $if$ ($ui$ == true)
    ''' 
    ''' Callback called by the UI manager when user clicks a button
    ''' Should do something meaninful wehn corresponding action is called.
    ''' 
    Public Sub OnCommand(commandId As String)
        Select commandId
            Case "Command1"
                MessageBox.Show(commandId)
                Return

            Case "Command2"
                MessageBox.Show(commandId)
                Return
			$endif$$if$ ($taskpaneANDui$ == true)
            Case "TogglePanel"
                TogglePanel()
                Return
        $endif$$if$ ($ui$ == true)End Select
    End Sub

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
                Return Application IsNot Nothing AndAlso Application.ActiveWindow IsNot Nothing
			$endif$$if$ ($taskpaneANDui$ == true)
            Case "TogglePanel"
                ' make panel enabled only if we have an open drawing.
                Return IsPanelEnabled()
            $endif$$if$ ($ui$ == true)Case Else
                Return True
        End Select
    End Function

    ''' 
    ''' Callback called by UI manager.
    ''' Should return if corresponding command (button) is pressed or not (makes sense for toggle buttons)
    ''' 
    Public Function IsCommandChecked(command As String) As Boolean
		$endif$$if$ ($taskpaneANDui$ == true)
        If command = "TogglePanel" Then
            Return IsPanelVisible()
        End If

        $endif$$if$($ui$ == true)Return False
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
    Public Function GetCommandIcon(command As String) As Icon
        Return DirectCast(My.Resources.ResourceManager.GetObject(command), Icon)
    End Function
	$endif$$if$ ($taskpane$ == true)
#Region "Panel"

    Private Sub TogglePanel()
        _panelManager.TogglePanel(Application.ActiveWindow)
    End Sub

    Private Function IsPanelEnabled() As Boolean
        Return Application IsNot Nothing AndAlso Application.ActiveWindow IsNot Nothing
    End Function

    Private Function IsPanelVisible() As Boolean
        Return Application IsNot Nothing AndAlso _panelManager.IsPanelOpened(Application.ActiveWindow)
    End Function

    Private _panelManager As PanelManager

#End Region
	$endif$
    Sub Startup(app As Object)
        Application = DirectCast(app, Microsoft.Office.Interop.Visio.Application)
		$if$ ($taskpane$ == true)
        _panelManager = New PanelManager(Me)
		$endif$$if$ ($ribbonANDcommandbars$ == true)Dim version = Integer.Parse(Application.Version, NumberStyles.AllowDecimalPoint)
        If (version < 14) Then
			$endif$$if$ ($commandbars$ == true)StartupCommandBars("$csprojectname$", New String() {"Command1", "Command2"$endif$$if$ ($commandbarsANDtaskpane$ == true), "TogglePanel"$endif$$if$ ($commandbars$ == true)})
        $endif$$if$ ($ribbonANDcommandbars$ == true)End If
    $endif$End Sub

    Sub Shutdown()
	    $if$ ($commandbars$ == true)ShutdownCommandBars()
        $endif$$if$ ($taskpane$ == true)_panelManager.Dispose()
    $endif$End Sub

	$if$ ($taskpaneORui$ == true)
    Sub UpdateUI()
        $endif$$if$ ($commandbars$ == true)UpdateCommandBars()
        $endif$$if$ ($ribbon$ == true)UpdateRibbon()
    $endif$$if$ ($taskpaneORui$ == true)
    End Sub
    $endif$
End Class