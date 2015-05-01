#region Copyright ©2005, Cristi Potlog - All Rights Reserved
/* ------------------------------------------------------------------- *
*                            Cristi Potlog                             *
*                  Copyright ©2005 - All Rights reserved               *
*                                                                      *
* THIS SOURCE CODE IS PROVIDED "AS IS" WITH NO WARRANTIES OF ANY KIND, *
* EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE        *
* WARRANTIES OF DESIGN, MERCHANTIBILITY AND FITNESS FOR A PARTICULAR   *
* PURPOSE, NONINFRINGEMENT, OR ARISING FROM A COURSE OF DEALING,       *
* USAGE OR TRADE PRACTICE.                                             *
*                                                                      *
* THIS COPYRIGHT NOTICE MAY NOT BE REMOVED FROM THIS FILE.             *
* ------------------------------------------------------------------- */
#endregion Copyright ©2005, Cristi Potlog - All Rights Reserved

#region References

using System;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;
using System.Windows.Forms.Design;

#endregion

namespace PanelAddinWizard
{

    #region Enums
    /// <summary>
    /// Represents possible styles of a wizard page.
    /// </summary>
    public enum WizardPageStyle
    {
        /// <summary>
        /// Represents a standard interior wizard page with a white banner at the top.
        /// </summary>
        Standard,
        /// <summary>
        /// Represents a welcome wizard page with white background and large logo on the left.
        /// </summary>
        Welcome,
        /// <summary>
        /// Represents a finish wizard page with white background,
        /// a large logo on the left and OK button.
        /// </summary>
        Finish,
        /// <summary>
        /// Represents a blank wizard page.
        /// </summary>
        Custom
    }
    #endregion

    /// <summary>
    /// Represents a wizard page control with basic layout functionality.
    /// </summary>
    [DefaultEvent("Click")]
    [Designer(typeof(WizardPageDesigner))]
    public class WizardPage : Panel
    {
        #region Consts
        private const int HeaderAreaHeight = 64;
        private const int HeaderGlyphSize = 32;
        private const int HeaderTextPadding = 12;
        private const int WelcomeGlyphWidth = 164;
        #endregion

        #region Fields
        private WizardPageStyle _style = WizardPageStyle.Standard;
        private string _title = String.Empty;
        private string _description = String.Empty;
        #endregion

        #region Constructor
        /// <summary>
        /// Creates a new instance of the <see cref="WizardPage"/> class.
        /// </summary>
        public WizardPage()
        {
            // reset control style to improove rendering (reduce flicker)
            SetStyle(ControlStyles.AllPaintingInWmPaint, true);
            SetStyle(ControlStyles.DoubleBuffer, true);
            SetStyle(ControlStyles.ResizeRedraw, true);
            SetStyle(ControlStyles.UserPaint, true);
        }
        #endregion

        #region Properties
        /// <summary>
        /// Gets or sets the style of the wizard page.
        /// </summary>
        [DefaultValue(WizardPageStyle.Standard)]
        [Category("Wizard")]
        [Description("Gets or sets the style of the wizard page.")]
        public WizardPageStyle Style
        {
            get
            {
                return _style;
            }
            set
            {
                if (_style != value)
                {
                    _style = value;
                    // get the parent wizard control
                    var parentWizard = Parent as Wizard;
                    if (parentWizard != null)
                    {
                        // check if page is selected
                        if (parentWizard.SelectedPage == this)
                        {
                            // reactivate the selected page (performs redraw too)
                            parentWizard.SelectedPage = this;
                        }
                    }
                    else
                    {
                        // just redraw the page
                        Invalidate();
                    }
                }
            }
        }

        /// <summary>
        /// Gets or sets the title of the wizard page.
        /// </summary>
        [DefaultValue("")]
        [Category("Wizard")]
        [Description("Gets or sets the title of the wizard page.")]
        public string Title
        {
            get
            {
                return _title;
            }
            set
            {
                if (value == null)
                {
                    value = String.Empty;
                }
                if (_title != value)
                {
                    _title = value;
                    Invalidate();
                }
            }
        }

        /// <summary>
        /// Gets or sets the description of the wizard page.
        /// </summary>
        [DefaultValue("")]
        [Category("Wizard")]
        [Description("Gets or sets the description of the wizard page.")]
        public string Description
        {
            get
            {
                return _description;
            }
            set
            {
                if (value == null)
                {
                    value = String.Empty;
                }
                if (_description != value)
                {
                    _description = value;
                    Invalidate();
                }
            }
        }
        #endregion

        #region Methods
        /// <summary>
        /// Provides custom drawing to the wizard page.
        /// </summary>
        protected override void OnPaint(PaintEventArgs e)
        {
            // raise paint event
            base.OnPaint(e);

            // check if custom style
            if (_style == WizardPageStyle.Custom)
            {
                // filter out
                return;
            }

            // init graphic resources
            Rectangle headerRect = ClientRectangle;
            Rectangle glyphRect = Rectangle.Empty;
            Rectangle titleRect = Rectangle.Empty;
            Rectangle descriptionRect = Rectangle.Empty;

            // determine text format
            StringFormat textFormat = StringFormat.GenericDefault;
            textFormat.LineAlignment = StringAlignment.Near;
            textFormat.Alignment = StringAlignment.Near;
            textFormat.Trimming = StringTrimming.EllipsisCharacter;

            var parentWizard = Parent as Wizard;

            switch (_style)
            {
                case WizardPageStyle.Standard:
                    // adjust height for header
                    headerRect.Height = HeaderAreaHeight;
                    // draw header border
                    ControlPaint.DrawBorder3D(e.Graphics, headerRect, Border3DStyle.Etched, Border3DSide.Bottom);
                    // adjust header rect not to overwrite the border
                    headerRect.Height -= SystemInformation.Border3DSize.Height;
                    // fill header with window color
                    e.Graphics.FillRectangle(SystemBrushes.Window, headerRect);

                    // determine header image regtangle
                    int headerPadding = ((HeaderAreaHeight - HeaderGlyphSize) / 2);
                    glyphRect.Location = new Point(Width - HeaderGlyphSize - headerPadding, headerPadding);
                    glyphRect.Size = new Size(HeaderGlyphSize, HeaderGlyphSize);

                    // determine the header content
                    Image headerImage = null;
                    Font headerFont = Font;
                    Font headerTitleFont = Font;
                    if (parentWizard != null)
                    {
                        // get content from parent wizard, if exists
                        headerImage = parentWizard.HeaderImage;
                        headerFont = parentWizard.HeaderFont;
                        headerTitleFont = parentWizard.HeaderTitleFont;
                    }

                    // check if we have an image
                    if (headerImage == null)
                    {
                        // display a focus rect as a place holder
                        ControlPaint.DrawFocusRectangle(e.Graphics, glyphRect);
                    }
                    else
                    {
                        // draw header image
                        e.Graphics.DrawImage(headerImage, glyphRect);
                    }

                    // determine title height
                    int headerTitleHeight = (int)Math.Ceiling(e.Graphics.MeasureString(_title, headerTitleFont, 0, textFormat).Height);

                    // calculate text sizes
                    titleRect.Location = new Point(HeaderTextPadding,
                                                    HeaderTextPadding);
                    titleRect.Size = new Size(glyphRect.Left - HeaderTextPadding,
                                                headerTitleHeight);
                    descriptionRect.Location = titleRect.Location;
                    descriptionRect.Y += headerTitleHeight + HeaderTextPadding / 2;
                    descriptionRect.Size = new Size(titleRect.Width,
                                                    HeaderAreaHeight - descriptionRect.Y);

                    // draw tilte text (single line, truncated with ellipsis)
                    e.Graphics.DrawString(_title,
                                            headerTitleFont,
                                            SystemBrushes.WindowText,
                                            titleRect,
                                            textFormat);
                    // draw description text (multiple lines if needed)
                    e.Graphics.DrawString(_description,
                                            headerFont,
                                            SystemBrushes.WindowText,
                                            descriptionRect,
                                            textFormat);
                    break;
                case WizardPageStyle.Welcome:
                case WizardPageStyle.Finish:
                    // fill whole page with window color
                    e.Graphics.FillRectangle(SystemBrushes.Window, headerRect);

                    // determine welcome image regtangle
                    glyphRect.Location = Point.Empty;
                    glyphRect.Size = new Size(WelcomeGlyphWidth, Height);

                    // determine the icon that should appear on the welcome page
                    Image welcomeImage = null;
                    Font welcomeFont = Font;
                    Font welcomeTitleFont = Font;
                    if (parentWizard != null)
                    {
                        welcomeImage = parentWizard.WelcomeImage;
                        welcomeFont = parentWizard.WelcomeFont;
                        welcomeTitleFont = parentWizard.WelcomeTitleFont;
                    }

                    // check if we have an image
                    if (welcomeImage == null)
                    {
                        // display a focus rect as a place holder
                        ControlPaint.DrawFocusRectangle(e.Graphics, glyphRect);
                    }
                    else
                    {
                        // draw welcome page image
                        e.Graphics.DrawImage(welcomeImage, glyphRect);
                    }

                    // calculate text sizes
                    titleRect.Location = new Point(WelcomeGlyphWidth + HeaderTextPadding,
                                                    HeaderTextPadding);
                    titleRect.Width = Width - titleRect.Left - HeaderTextPadding;
                    // determine title height
                    int welcomeTitleHeight = (int)Math.Ceiling(e.Graphics.MeasureString(_title, welcomeTitleFont, titleRect.Width, textFormat).Height);
                    descriptionRect.Location = titleRect.Location;
                    descriptionRect.Y += welcomeTitleHeight + HeaderTextPadding;
                    descriptionRect.Size = new Size(Width - descriptionRect.Left - HeaderTextPadding,
                                                    Height - descriptionRect.Y);

                    // draw tilte text (multiple lines if needed)
                    e.Graphics.DrawString(_title,
                                            welcomeTitleFont,
                                            SystemBrushes.WindowText,
                                            titleRect,
                                            textFormat);
                    // draw description text (multiple lines if needed)
                    e.Graphics.DrawString(_description,
                                            welcomeFont,
                                            SystemBrushes.WindowText,
                                            descriptionRect,
                                            textFormat);
                    break;
            }
        }
        #endregion

        #region Inner classes
        /// <summary>
        /// This is a designer for the Banner.
        /// This designer locks the control vertical sizing.
        /// </summary>
        internal class WizardPageDesigner : ParentControlDesigner
        {
            /// <summary>
            /// Gets the selection rules that indicate the movement capabilities of a component.
            /// </summary>
            public override SelectionRules SelectionRules
            {
                get
                {
                    // lock the control
                    return SelectionRules.Visible | SelectionRules.Locked;
                }
            }
        }
        #endregion
    }
}
