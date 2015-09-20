Imports System
Imports System.Collections.Generic


''' 
''' Manages the list of all installed panels
''' 
Public Class PanelManager
    Implements IDisposable

    Public Sub New()
        $if$ ($uiCallbacks$ == true)AddHandler Globals.ThisAddIn.Application.DocumentCreated, AddressOf OnDocumentListChanged
        AddHandler Globals.ThisAddIn.Application.DocumentOpened, AddressOf OnDocumentListChanged
        AddHandler Globals.ThisAddIn.Application.BeforeDocumentClose, AddressOf OnDocumentListChanged
    $endif$End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        $if$ ($uiCallbacks$ == true)RemoveHandler Globals.ThisAddIn.Application.DocumentCreated, AddressOf OnDocumentListChanged
        RemoveHandler Globals.ThisAddIn.Application.DocumentOpened, AddressOf OnDocumentListChanged
        RemoveHandler Globals.ThisAddIn.Application.BeforeDocumentClose, AddressOf OnDocumentListChanged
    $endif$End Sub
    $if$ ($uiCallbacks$ == true)
    Private Sub OnDocumentListChanged(doc As Microsoft.Office.Interop.Visio.Document)
        Globals.ThisAddIn.UpdateUI()
    End Sub
    $endif$
    Private ReadOnly _panelFrames As New Dictionary(Of Integer, PanelFrame)()

    Private Function FindWindowPanelFrame(window As Microsoft.Office.Interop.Visio.Window) As PanelFrame
        If window Is Nothing Then
            Return Nothing
        End If

        Return If(_panelFrames.ContainsKey(window.ID), _panelFrames(window.ID), Nothing)
    End Function

    ''' 
    ''' Shows or hides panel for the given Visio window.
    ''' 
    Public Sub TogglePanel(window As Microsoft.Office.Interop.Visio.Window)
        If window Is Nothing Then
            Return
        End If

        Dim panelFrame = FindWindowPanelFrame(window)
        If panelFrame Is Nothing Then
            panelFrame = New PanelFrame(New TheForm(window))
            panelFrame.CreateWindow(window)

            AddHandler panelFrame.PanelFrameClosed, AddressOf OnPanelFrameClosed
            _panelFrames.Add(window.ID, panelFrame)
        Else
            panelFrame.DestroyWindow()
            _panelFrames.Remove(window.ID)
        End If
        $if$ ($uiCallbacks$ == true)
        Globals.ThisAddIn.UpdateUI()
    $endif$End Sub

    Private Sub OnPanelFrameClosed(window As Microsoft.Office.Interop.Visio.Window)
        _panelFrames.Remove(window.ID)
        $if$ ($uiCallbacks$ == true)
        Globals.ThisAddIn.UpdateUI()
    $endif$End Sub

        ''' 
        ''' Returns true if panel is opened in the given Visio diagram window.
        ''' 
    Public Function IsPanelOpened(window As Microsoft.Office.Interop.Visio.Window) As Boolean
        Return FindWindowPanelFrame(window) IsNot Nothing
    End Function
End Class