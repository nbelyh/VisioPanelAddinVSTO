Imports System
Imports System.Collections.Generic


''' 
''' Manages the list of all installed panels
''' 
Public Class PanelManager
    Implements IDisposable
    Private ReadOnly _addin As Addin

    Public Sub New(addin As Addin)
        _addin = addin

        AddHandler _addin.Application.DocumentCreated, AddressOf OnDocumentListChanged
        AddHandler _addin.Application.DocumentOpened, AddressOf OnDocumentListChanged
        AddHandler _addin.Application.BeforeDocumentClose, AddressOf OnDocumentListChanged
    End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        RemoveHandler _addin.Application.DocumentCreated, AddressOf OnDocumentListChanged
        RemoveHandler _addin.Application.DocumentOpened, AddressOf OnDocumentListChanged
        RemoveHandler _addin.Application.BeforeDocumentClose, AddressOf OnDocumentListChanged
    End Sub

    Private Sub OnDocumentListChanged(doc As Microsoft.Office.Interop.Visio.Document)
        _addin.UpdateUI()
    End Sub

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

        _addin.UpdateUI()
    End Sub

    Private Sub OnPanelFrameClosed(window As Microsoft.Office.Interop.Visio.Window)
        _panelFrames.Remove(window.ID)

        _addin.UpdateUI()
    End Sub

    ''' 
    ''' Returns true if panel is opened in the given Visio diagram window.
    ''' 
    Public Function IsPanelOpened(window As Microsoft.Office.Interop.Visio.Window) As Boolean
        Return FindWindowPanelFrame(window) IsNot Nothing
    End Function
End Class