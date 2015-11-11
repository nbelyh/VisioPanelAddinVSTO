using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using stdole;
using System.Windows.Forms;
using System.Drawing.Imaging;

namespace $csprojectname$
{
    /// <summary>
    /// User interface manager for Visio 2007 and below 
    /// (works also with latest versions, but makes no sense there, as there is ribbon)
    /// Creates and controls a custom command bar with buttons
    /// </summary>
    /// 
    public partial class AddinUI
    {
        private string _toolbarName;

        private readonly List<string> _commands =
            new List<string>();

        /// <summary>
        /// Constructs the UI manager
        /// </summary>
        /// <param name="toolbarName">The name of the toolbar to install</param>
        /// <param name="commands">The list of buttons to add (list of button ids)</param>
        public void StartupCommandBars(string toolbarName, IEnumerable<string> commands)
        {
            _toolbarName = toolbarName;
            _commands.AddRange(commands);

            ThisAddIn.Application.VisioIsIdle += ApplicationIdle;
            UpdateCommandBars();
        }

        public void ShutdownCommandBars()
        {
            ThisAddIn.Application.VisioIsIdle -= ApplicationIdle;
        }

        private bool _updateRequest;

        internal void UpdateCommandBars()
        {
            _updateRequest = true;
        }

        void ApplicationIdle(object app)
        {
            if (!_updateRequest)
                return;

            _updateRequest = false;

            UpdateToolbar();
        }

        private readonly Dictionary<string, CommandBarButton> _buttons =
            new Dictionary<string, CommandBarButton>();

        private static CommandBar FindCommandBar(CommandBars cbs, string name)
        {
            return cbs.Cast<CommandBar>().FirstOrDefault(cb => cb.Name == name);
        }

        private void UpdateToolbar()
        {
            var cbs = (CommandBars)ThisAddIn.Application.CommandBars;

            var cb = FindCommandBar(cbs, _toolbarName) ?? cbs.Add(_toolbarName);
            cb.Visible = true;

            foreach (var id in _commands)
                InstallButton(cb, id);
        }

        private void InstallButton(CommandBar cb, string id)
        {
            CommandBarButton thisButton;
            _buttons.TryGetValue(id, out thisButton);

            // Recreate the button, otherwise it's state may be broken in Visio in some cases
            if (thisButton != null)
            {
                thisButton.Click -= CommandBarButtonClicked;
                _buttons.Remove(id);
                Marshal.ReleaseComObject(thisButton);
            }

            var button = (CommandBarButton)cb.FindControl(Tag: id) ??
                (CommandBarButton)cb.Controls.Add(MsoControlType.msoControlButton);

            button.Enabled = ThisAddIn.IsCommandEnabled(id);

            var checkState = ThisAddIn.IsCommandChecked(id);
            button.State = checkState ? MsoButtonState.msoButtonDown : MsoButtonState.msoButtonUp;

            button.Tag = id;
            button.Caption = ThisAddIn.GetCommandLabel(id);
            SetCommandBarButtonImage(button, id);

            button.Click += CommandBarButtonClicked;

            _buttons.Add(id, button);
        }

        private void CommandBarButtonClicked(CommandBarButton ctrl, ref bool cancelDefault)
        {
            ThisAddIn.OnCommand(ctrl.Tag);
        }

        private void SetCommandBarButtonImage(CommandBarButton button, string id)
        {
            var image = ThisAddIn.GetCommandBitmap(id + "_sm");
            if (image == null)
                return;

            Bitmap picture, mask;
            BitmapToPictureAndMask(image, out picture, out mask);

            button.Picture = PictureConvert.ImageToPictureDisp(picture);
            button.Mask = PictureConvert.ImageToPictureDisp(mask);
        }

        static public void BitmapToPictureAndMask(Bitmap bm, out Bitmap picture, out Bitmap mask)
        {
            var w = bm.Width;
            var h = bm.Height;

            var pictureData = new byte[3 * w * h];
            var maskData = new byte[3 * w * h];

            var bmBits = bm.LockBits(new Rectangle(0, 0, w, h), ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);
            var bits = new byte[4 * w * h];
            Marshal.Copy(bmBits.Scan0, bits, 0, 4 * w * h);
            bm.UnlockBits(bmBits);

            for (var y = 0; y < h; ++y)
            {
                for (var x = 0; x < w; ++x)
                {
                    var srcIdx = (x + y * w) * 4;
                    var dstIdx = (x + y * w) * 3;

                    pictureData[dstIdx + 0] = bits[srcIdx + 0];
                    pictureData[dstIdx + 1] = bits[srcIdx + 1];
                    pictureData[dstIdx + 2] = bits[srcIdx + 2];

                    var t = (bits[srcIdx + 3] < 128) ? (byte)255 : (byte)0;

                    maskData[dstIdx + 0] = t;
                    maskData[dstIdx + 1] = t;
                    maskData[dstIdx + 2] = t;
                }
            }

            var rect = new Rectangle(0, 0, w, h);

            picture = new Bitmap(w, h, PixelFormat.Format24bppRgb);
            var pictureBits = picture.LockBits(rect, ImageLockMode.WriteOnly, picture.PixelFormat);
            Marshal.Copy(pictureData, 0, pictureBits.Scan0, w * h * 3);
            picture.UnlockBits(pictureBits);

            mask = new Bitmap(w, h, PixelFormat.Format24bppRgb);
            var maskBits = mask.LockBits(rect, ImageLockMode.WriteOnly, picture.PixelFormat);
            Marshal.Copy(maskData, 0, maskBits.Scan0, w * h * 3);
            mask.UnlockBits(maskBits);
        }

        [System.ComponentModel.DesignerCategory("")]
        class PictureConvert : AxHost
        {
            private PictureConvert() : base("") { }

            static public IPictureDisp ImageToPictureDisp(Bitmap image)
            {
                return (IPictureDisp)GetIPictureDispFromPicture(image);
            }
        }
    }
}
