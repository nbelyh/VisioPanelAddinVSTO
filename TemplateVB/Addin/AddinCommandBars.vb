
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.Linq
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports Microsoft.Office.Core
Imports stdole

Partial Public Class AddinUI
    Private _toolbarName As String

    Private ReadOnly _commands As New List(Of String)()

    ''' 
    ''' Constructs the UI manager
    ''' 
    Public Sub StartupCommandBars(toolbarName As String, commands As IEnumerable(Of String))
        _toolbarName = toolbarName
        _commands.AddRange(commands)

        AddHandler ThisAddIn.Application.VisioIsIdle, AddressOf ApplicationIdle
        UpdateCommandBars()
    End Sub

    Public Sub ShutdownCommandBars()
        RemoveHandler ThisAddIn.Application.VisioIsIdle, AddressOf ApplicationIdle
    End Sub

    Private _updateRequest As Boolean

    Sub UpdateCommandBars()
        _updateRequest = True
    End Sub

    Private Sub ApplicationIdle(app As Object)
        If Not _updateRequest Then
            Return
        End If

        _updateRequest = False

        UpdateToolbar()
    End Sub

    Private ReadOnly _buttons As New Dictionary(Of String, CommandBarButton)()

    Private Shared Function FindCommandBar(cbs As CommandBars, name As String) As CommandBar
        Return cbs.Cast(Of CommandBar)().FirstOrDefault(Function(cb) cb.Name = name)
    End Function

    Private Sub UpdateToolbar()
        Dim cbs = DirectCast(ThisAddIn.Application.CommandBars, CommandBars)

        Dim cb = If(FindCommandBar(cbs, _toolbarName), cbs.Add(_toolbarName))
        cb.Visible = True

        For Each id In _commands
            InstallButton(cb, id)
        Next
    End Sub

    Private Sub InstallButton(cb As CommandBar, id As String)
        Dim thisButton As CommandBarButton = Nothing
        _buttons.TryGetValue(id, thisButton)

        ' Recreate the button, otherwise it's state may be broken in Visio in some cases
        If thisButton IsNot Nothing Then
            RemoveHandler thisButton.Click, AddressOf CommandBarButtonClicked
            _buttons.Remove(id)
            Marshal.ReleaseComObject(thisButton)
        End If

        Dim button = If(DirectCast(cb.FindControl(Tag:=id), CommandBarButton), DirectCast(cb.Controls.Add(MsoControlType.msoControlButton), CommandBarButton))

        button.Enabled = ThisAddIn.IsCommandEnabled(id)

        Dim checkState = ThisAddIn.IsCommandChecked(id)
        button.State = If(checkState, MsoButtonState.msoButtonDown, MsoButtonState.msoButtonUp)

        button.Tag = id
        button.Caption = ThisAddIn.GetCommandLabel(id)
        SetCommandBarButtonImage(button, id)

        AddHandler button.Click, AddressOf CommandBarButtonClicked

        _buttons.Add(id, button)
    End Sub

    Private Sub CommandBarButtonClicked(ctrl As CommandBarButton, ByRef cancelDefault As Boolean)
        ThisAddIn.OnCommand(ctrl.Tag)
    End Sub

    Private Sub SetCommandBarButtonImage(button As CommandBarButton, id As String)
        Dim image = ThisAddIn.GetCommandBitmap(id & "_sm")
        If image Is Nothing Then
            Return
        End If

        Dim picture As Bitmap = Nothing
        Dim mask As Bitmap = Nothing
        BitmapToPictureAndMask(image, picture, mask)

        button.Picture = PictureConvert.ImageToPictureDisp(picture)
        button.Mask = PictureConvert.ImageToPictureDisp(mask)
    End Sub

    Public Shared Sub BitmapToPictureAndMask(bm As Bitmap, ByRef picture As Bitmap, ByRef mask As Bitmap)
        Dim w = bm.Width
        Dim h = bm.Height

        Dim pictureData = New Byte(3 * w * h - 1) {}
        Dim maskData = New Byte(3 * w * h - 1) {}

        Dim bmBits = bm.LockBits(New Rectangle(0, 0, w, h), ImageLockMode.[ReadOnly], PixelFormat.Format32bppArgb)
        Dim bits = New Byte(4 * w * h - 1) {}
        Marshal.Copy(bmBits.Scan0, bits, 0, 4 * w * h)
        bm.UnlockBits(bmBits)

        For y = 0 To h - 1
            For x = 0 To w - 1
                Dim srcIdx = (x + y * w) * 4
                Dim dstIdx = (x + y * w) * 3

                pictureData(dstIdx + 0) = bits(srcIdx + 0)
                pictureData(dstIdx + 1) = bits(srcIdx + 1)
                pictureData(dstIdx + 2) = bits(srcIdx + 2)

                Dim t = If((bits(srcIdx + 3) < 128), CByte(255), CByte(0))

                maskData(dstIdx + 0) = t
                maskData(dstIdx + 1) = t
                maskData(dstIdx + 2) = t
            Next
        Next

        Dim rect = New Rectangle(0, 0, w, h)

        picture = New Bitmap(w, h, PixelFormat.Format24bppRgb)
        Dim pictureBits = picture.LockBits(rect, ImageLockMode.[WriteOnly], picture.PixelFormat)
        Marshal.Copy(pictureData, 0, pictureBits.Scan0, w * h * 3)
        picture.UnlockBits(pictureBits)

        mask = New Bitmap(w, h, PixelFormat.Format24bppRgb)
        Dim maskBits = mask.LockBits(rect, ImageLockMode.[WriteOnly], picture.PixelFormat)
        Marshal.Copy(maskData, 0, maskBits.Scan0, w * h * 3)
        mask.UnlockBits(maskBits)
    End Sub

    Private Class PictureConvert
        Inherits AxHost
        Private Sub New()
            MyBase.New("")
        End Sub

        Public Shared Function ImageToPictureDisp(image As Bitmap) As IPictureDisp
            Return DirectCast(GetIPictureDispFromPicture(image), IPictureDisp)
        End Function
    End Class
End Class