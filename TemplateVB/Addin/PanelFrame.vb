Imports System
Imports System.Diagnostics
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Visio


''' 
''' Integrates a winform in Visio.
''' Creates an anchor window for the given diagram window, and installs the specified form as a child in that panel.
''' 
Public NotInheritable Class PanelFrame
    Implements IVisEventProc
    Private Const AddonWindowMergeId As String = "$mergeguid$"

#Region "WIN API Declares"

    <DllImport("user32.dll", EntryPoint:="SetWindowLong", CharSet:=CharSet.Auto, CallingConvention:=CallingConvention.Winapi)> _
    Private Shared Function SetWindowLong(hWnd As IntPtr, index As Integer, newLong As Integer) As Integer
    End Function

    <DllImport("user32.dll", EntryPoint:="SetParent", CharSet:=CharSet.Auto, CallingConvention:=CallingConvention.Winapi)> _
    Private Shared Function SetParent(hWndChild As IntPtr, hWndNewParent As IntPtr) As Integer
    End Function

    <DllImport("user32.dll", EntryPoint:="GetWindowRect", CharSet:=CharSet.Auto, CallingConvention:=CallingConvention.Winapi)> _
    Private Shared Function GetWindowRect(hWnd As IntPtr, ByRef lpRect As RECT) As Integer
    End Function

    <DllImport("user32.dll", EntryPoint:="SetWindowPos", CharSet:=CharSet.Auto, CallingConvention:=CallingConvention.Winapi)> _
    Private Shared Function SetWindowPos(hWnd As IntPtr, hWndInsertAfter As IntPtr, x As Integer, y As Integer, cx As Integer, cy As Integer, _
                                         wFlags As Integer) As Integer
    End Function

    Private Const GWL_STYLE As Integer = (-16)
    Private Const WS_CHILD As Integer = &H40000000
    Private Const WS_OVERLAPPED As Integer = &H0
    Private Const SWP_NOCOPYBITS As Integer = &H100
    Private Const SWP_NOMOVE As Integer = &H2
    Private Const SWP_NOZORDER As Integer = &H4
    Private Const GW_CHILD As Integer = 5
    Private Const GW_HWNDNEXT As Integer = 2

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Auto)> _
    Private Structure RECT
        Public left As Integer
        Public top As Integer
        Public right As Integer
        Public bottom As Integer
    End Structure

#End Region

#Region "fields"

    Private _visioWindow As Window
    Private _form As Form

#End Region

    ''' 
    ''' The event is triggered when user closes the panel using "x" button
    ''' 
    Public Delegate Sub PanelFrameClosedEventHandler(window As Window)
    Public Event PanelFrameClosed As PanelFrameClosedEventHandler

    ''' 
    ''' Constructs a new panel frame.
    ''' 
    Public Sub New(form As Form)
        _form = form
    End Sub

#Region "methods"

    ''' Destroys the panel frame along with the form.
    ''' 
    Public Sub DestroyWindow()
        Try
            If _visioWindow IsNot Nothing AndAlso _form IsNot Nothing Then
                Dim childWindowHandle = _form.Handle

                _form.Hide()

                SetWindowLong(childWindowHandle, GWL_STYLE, WS_OVERLAPPED)
                SetParent(childWindowHandle, CType(0, IntPtr))

                _visioWindow.Close()
                _visioWindow = Nothing
            End If

            If _form IsNot Nothing Then
                _form.Close()
                _form.Dispose()
                _form = Nothing
            End If
            ' ReSharper disable once EmptyGeneralCatchClause : ignore all errors on exit
        Catch
        End Try
    End Sub

    ''' Install the panel into given window (actually creates the form and shows it)
    ''' 
    Public Function CreateWindow(visioParentWindow As Window) As Window
        Dim retVal As Window = Nothing

        Try
            If visioParentWindow Is Nothing Then
                Return Nothing
            End If

            If _form IsNot Nothing Then
                Dim childWindowHandle As IntPtr = _form.Handle

                _visioWindow = visioParentWindow.Windows.Add(_form.Text, CInt(VisWindowStates.visWSDockedRight) Or CInt(VisWindowStates.visWSAnchorMerged), VisWinTypes.visAnchorBarAddon, 0, 0, 300, _
                                                             300, AddonWindowMergeId, String.Empty, 0)

                AddHandler _visioWindow.BeforeWindowClosed, AddressOf OnBeforeWindowClosed

                _visioWindow.Visible = False

                Dim parentWindowHandle = CType(_visioWindow.WindowHandle32, IntPtr)

                SetWindowLong(childWindowHandle, GWL_STYLE, WS_CHILD)
                SetParent(childWindowHandle, parentWindowHandle)

                _form.Show()

                JiggleWindow(parentWindowHandle)

                _visioWindow.Visible = True
                _visioWindow.Activate()

                retVal = _visioWindow
            End If
        Catch ex As Exception
            Debug.Write(ex.Message)
        End Try

        Return retVal
    End Function

    Private Shared Sub JiggleWindow(handle As IntPtr)
        Dim lpRect = New RECT()
        GetWindowRect(handle, lpRect)

        Dim l = lpRect.left
        Dim T = lpRect.top
        Dim w = lpRect.right - lpRect.left
        Dim h = lpRect.bottom - lpRect.top

        Const flags As Integer = SWP_NOCOPYBITS Or SWP_NOMOVE Or SWP_NOZORDER
        SetWindowPos(handle, New IntPtr(0), l, T, w, h + 1, _
                     flags)
        SetWindowPos(handle, New IntPtr(0), l, T, w, h, _
                     flags)
    End Sub

#End Region

    Private Function IVisEventProc_VisEventProc(nEventCode As Short, pSourceObj As Object, nEventId As Integer, nEventSeqNum As Integer, pSubjectObj As Object, vMoreInfo As Object) As Object Implements IVisEventProc.VisEventProc
        Dim returnValue As Object = False

        Try
            Dim subjectWindow = TryCast(pSubjectObj, Window)
            Select Case nEventCode
                Case (CShort(VisEventCodes.visEvtDel) + CShort(VisEventCodes.visEvtWindow))
                    If True Then
                        OnBeforeWindowClosed(subjectWindow)
                        Exit Select
                    End If
            End Select
        Catch ex As Exception
            Debug.Write(ex.Message)
        End Try

        Return returnValue
    End Function

    Private Sub OnBeforeWindowClosed(visioWindow As Window)
        RaiseEvent PanelFrameClosed(_visioWindow.ParentWindow)

        DestroyWindow()
    End Sub
End Class