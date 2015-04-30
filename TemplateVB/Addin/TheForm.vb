Imports System.Windows.Forms
Imports Visio = Microsoft.Office.Interop.Visio

Public Class TheForm

    Private ReadOnly _window As Visio.Window

    Public Sub New(window As Visio.Window)
        _window = window
        InitializeComponent()
    End Sub

    Private Sub button1_Click(sender As Object, e As EventArgs) Handles button1.Click
        MessageBox.Show(_window.Document.Name)
    End Sub

End Class