using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Office.Interop.Visio;

namespace $csprojectname$
{
    /// <summary>
    /// Integrates a winform in Visio.
    /// Creates an anchor window for the given diagram window, and installs the specified form as a child in that panel.
    /// </summary>
    /// 
    public sealed class PanelFrame : IVisEventProc
    {
        private const string AddonWindowMergeId = "$mergeguid$";

        #region WIN API Declares

        [DllImport("user32.dll", EntryPoint = "SetWindowLong", CharSet = CharSet.Auto,
                    CallingConvention = CallingConvention.Winapi)]
        private static extern int SetWindowLong(IntPtr hWnd, int index, int newLong);

        [DllImport("user32.dll", EntryPoint = "SetParent", CharSet = CharSet.Auto,
                    CallingConvention = CallingConvention.Winapi)]
        private static extern int SetParent(IntPtr hWndChild, IntPtr hWndNewParent);

        [DllImport("user32.dll", EntryPoint = "GetWindowRect", CharSet = CharSet.Auto,
                    CallingConvention = CallingConvention.Winapi)]
        private static extern int GetWindowRect(IntPtr hWnd, ref RECT lpRect);

        [DllImport("user32.dll", EntryPoint = "SetWindowPos", CharSet = CharSet.Auto,
                    CallingConvention = CallingConvention.Winapi)]
        private static extern int SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter,
                    int x, int y, int cx, int cy, int wFlags);

        private const int GWL_STYLE = (-16);
        private const int WS_CHILD = 0x40000000;
        private const int WS_OVERLAPPED = 0x00000000;
        private const int SWP_NOCOPYBITS = 0x100;
        private const int SWP_NOMOVE = 0x2;
        private const int SWP_NOZORDER = 0x4;
        private const int GW_CHILD = 5;
        private const int GW_HWNDNEXT = 2;

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        private struct RECT
        {
            public int left;
            public int top;
            public int right;
            public int bottom;
        }

        #endregion

        #region fields

        private Window _visioWindow;
        private Form _form;

        #endregion

        /// <summary>
        /// The event is triggered when user closes the panel using "x" button
        /// </summary>
        /// <param name="window">The parent diagram window for which the panel was closed.</param>
        /// 
        public delegate void PanelFrameClosedEventHandler(Window window);
        public event PanelFrameClosedEventHandler PanelFrameClosed;

        /// <summary>
        /// Constructs a new panel frame.
        /// </summary>
        /// <param name="form">The form to install</param>
        public PanelFrame(Form form)
        {
            _form = form;
        }

        #region methods

        /// <summary>
        /// Destroys the panel frame along with the form.
        /// </summary>
        public void DestroyWindow()
        {
            try
            {
                if (_visioWindow != null && _form != null)
                {
                    var childWindowHandle = _form.Handle;

                    _form.Hide();

                    SetWindowLong(childWindowHandle, GWL_STYLE, WS_OVERLAPPED);
                    SetParent(childWindowHandle, (IntPtr)0);

                    _visioWindow.Close();
                    _visioWindow = null;
                }

                if (_form != null)
                {
                    _form.Close();
                    _form.Dispose();
                    _form = null;
                }
            }
            // ReSharper disable once EmptyGeneralCatchClause : ignore all errors on exit
            catch
            {
            }
        }

        /// <summary>
        /// Install the panel into given window (actually creates the form and shows it)
        /// </summary>
        /// <param name="visioParentWindow">The parent Visio window where the panel should be installed to.</param>
        /// <returns></returns>
        public Window CreateWindow(Window visioParentWindow)
        {
            Window retVal = null;

            try
            {
                if (visioParentWindow == null)
                    return null;

                if (_form != null)
                {
                    IntPtr childWindowHandle = _form.Handle;

                    _visioWindow = visioParentWindow.Windows.Add(
                        _form.Text,
                        (int)VisWindowStates.visWSDockedRight | (int)VisWindowStates.visWSAnchorMerged,
                        VisWinTypes.visAnchorBarAddon,
                        0,
                        0,
                        300,
                        300,
                        AddonWindowMergeId,
                        string.Empty,
                        0);

                    _visioWindow.BeforeWindowClosed += OnBeforeWindowClosed;

                    _visioWindow.Visible = false;

                    var parentWindowHandle = (IntPtr)_visioWindow.WindowHandle32;

                    SetWindowLong(childWindowHandle, GWL_STYLE, WS_CHILD);
                    SetParent(childWindowHandle, parentWindowHandle);

                    _form.Show();

                    JiggleWindow(parentWindowHandle);

                    _visioWindow.Visible = true;
                    _visioWindow.Activate();

                    retVal = _visioWindow;
                }
            }
            catch (Exception ex)
            {
                Debug.Write(ex.Message);
            }

            return retVal;
        }

        private static void JiggleWindow(
            IntPtr handle)
        {
            var lpRect = new RECT();
            GetWindowRect(handle, ref lpRect);

            var l = lpRect.left;
            var T = lpRect.top;
            var w = lpRect.right - lpRect.left;
            var h = lpRect.bottom - lpRect.top;

            const int flags = SWP_NOCOPYBITS | SWP_NOMOVE | SWP_NOZORDER;
            SetWindowPos(handle, new IntPtr(0), l, T, w, h + 1, flags);
            SetWindowPos(handle, new IntPtr(0), l, T, w, h, flags);
        }

        #endregion

        object IVisEventProc.VisEventProc(short nEventCode, object pSourceObj, int nEventId, int nEventSeqNum, object pSubjectObj, object vMoreInfo)
        {
            object returnValue = false;

            try
            {
                var subjectWindow = pSubjectObj as Window;
                switch (nEventCode)
                {
                    case ((short)VisEventCodes.visEvtDel + (short)VisEventCodes.visEvtWindow):
                        {
                            OnBeforeWindowClosed(subjectWindow);
                            break;
                        }
                }
            }
            catch (Exception ex)
            {
                Debug.Write(ex.Message);
            }

            return returnValue;
        }

        private void OnBeforeWindowClosed(Window visioWindow)
        {
            if (PanelFrameClosed != null)
                PanelFrameClosed(_visioWindow.ParentWindow);

            DestroyWindow();
        }
    }
}