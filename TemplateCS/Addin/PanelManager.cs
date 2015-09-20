using System;
using System.Collections.Generic;
using Visio = Microsoft.Office.Interop.Visio;

namespace $csprojectname$
{
    /// <summary>
    /// Manages the list of all installed panels
    /// </summary>
    /// 
    public class PanelManager : IDisposable
    {
        public PanelManager()
        {
            $if$ ($uiCallbacks$ == true)Globals.ThisAddIn.Application.DocumentCreated += OnDocumentListChanged;
            Globals.ThisAddIn.Application.DocumentOpened += OnDocumentListChanged;
            Globals.ThisAddIn.Application.BeforeDocumentClose += OnDocumentListChanged;
        $endif$}

        public void Dispose()
        {
            $if$ ($uiCallbacks$ == true)Globals.ThisAddIn.Application.DocumentCreated -= OnDocumentListChanged;
            Globals.ThisAddIn.Application.DocumentOpened -= OnDocumentListChanged;
            Globals.ThisAddIn.Application.BeforeDocumentClose -= OnDocumentListChanged;
        $endif$}
        $if$ ($uiCallbacks$ == true)
        private void OnDocumentListChanged(Visio.Document doc)
        {
            Globals.ThisAddIn.UpdateUI();
        }
        $endif$
        private readonly Dictionary<int, PanelFrame> _panelFrames =
            new Dictionary<int, PanelFrame>();

        PanelFrame FindWindowPanelFrame(Visio.Window window)
        {
            if (window == null)
                return null;

            return _panelFrames.ContainsKey(window.ID) ? _panelFrames[window.ID] : null;
        }

        /// <summary>
        /// Shows or hides panel for the given Visio window.
        /// </summary>
        /// <param name="window">Target Visio diagram window where to show/hide the panel</param>
        public void TogglePanel(Visio.Window window)
        {
            if (window == null)
                return;

            var panelFrame = FindWindowPanelFrame(window);
            if (panelFrame == null)
            {
                panelFrame = new PanelFrame(new TheForm(window));
                panelFrame.CreateWindow(window);

                panelFrame.PanelFrameClosed += OnPanelFrameClosed;
                _panelFrames.Add(window.ID, panelFrame);
            }
            else
            {
                panelFrame.DestroyWindow();
                _panelFrames.Remove(window.ID);
            }
            $if$ ($uiCallbacks$ == true)
            Globals.ThisAddIn.UpdateUI();
        $endif$}

        private void OnPanelFrameClosed(Visio.Window window)
        {
            _panelFrames.Remove(window.ID);
            $if$ ($uiCallbacks$ == true)
            Globals.ThisAddIn.UpdateUI();
        $endif$}

        /// <summary>
        /// Returns true if panel is opened in the given Visio diagram window.
        /// </summary>
        /// <param name="window">Visio diagram window</param>
        /// <returns></returns>
        public bool IsPanelOpened(Visio.Window window)
        {
            return FindWindowPanelFrame(window) != null;
        }
    }
}
