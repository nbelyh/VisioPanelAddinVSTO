#region Copyright �2005, Cristi Potlog - All Rights Reserved
/* ------------------------------------------------------------------- *
*                            Cristi Potlog                             *
*                  Copyright �2005 - All Rights reserved               *
*                                                                      *
* THIS SOURCE CODE IS PROVIDED "AS IS" WITH NO WARRANTIES OF ANY KIND, *
* EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE        *
* WARRANTIES OF DESIGN, MERCHANTIBILITY AND FITNESS FOR A PARTICULAR   *
* PURPOSE, NONINFRINGEMENT, OR ARISING FROM A COURSE OF DEALING,       *
* USAGE OR TRADE PRACTICE.                                             *
*                                                                      *
* THIS COPYRIGHT NOTICE MAY NOT BE REMOVED FROM THIS FILE.             *
* ------------------------------------------------------------------- */
#endregion Copyright �2005, Cristi Potlog - All Rights Reserved

#region References

using System;
using System.ComponentModel;
using System.ComponentModel.Design;
using System.Drawing;
using System.Windows.Forms;
using System.Windows.Forms.Design;
using PanelAddinWizard;

#endregion

namespace PanelAddinWizard
{
	/// <summary>
	/// Represents an extentable wizard control with basic page navigation functionality.
	/// </summary>
	[Designer(typeof(WizardDesigner))]
	public class Wizard : UserControl
	{
		private int FooterAreaHeight;
		private Point _offsetCancel;
		private Point _offsetNext;
		private Point _offsetBack;
		
		private WizardPage _selectedPage;
		private readonly WizardPagesCollection _pages;
		private Image _headerImage;
		private Image _welcomeImage;
		private Font _headerFont;
		private Font _headerTitleFont;
		private Font _welcomeFont;
		private Font _welcomeTitleFont;

		#region Designer generated code
		private System.Windows.Forms.Button buttonCancel;
		private System.Windows.Forms.Button buttonNext;
		private System.Windows.Forms.Button buttonBack;
		private System.Windows.Forms.Button buttonHelp;
		/// <summary> 
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;
		#endregion

		#region Constructor&Dispose
		/// <summary>
		/// Creates a new instance of the <see cref="Wizard"/> class.
		/// </summary>
		public Wizard()
		{
			// call required by designer
			InitializeComponent();

			// reset control style to improove rendering (reduce flicker)
			SetStyle(ControlStyles.AllPaintingInWmPaint, true); 
			SetStyle(ControlStyles.DoubleBuffer, true);
			SetStyle(ControlStyles.ResizeRedraw, true);
			SetStyle(ControlStyles.UserPaint, true);

			// reset dock style
			base.Dock = DockStyle.Fill;

			// init pages collection
			_pages = new WizardPagesCollection(this);
		}

		/// <summary> 
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose(bool disposing)
		{
			if (disposing)
			{
				if (components != null)
				{
					components.Dispose();
				}
			}
			base.Dispose(disposing);
		}

		#endregion

		#region Designer generated code
		/// <summary> 
		/// Required method for Designer support - do not modify 
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
            this.buttonCancel = new System.Windows.Forms.Button();
            this.buttonNext = new System.Windows.Forms.Button();
            this.buttonBack = new System.Windows.Forms.Button();
            this.buttonHelp = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // buttonCancel
            // 
            this.buttonCancel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.buttonCancel.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.buttonCancel.Location = new System.Drawing.Point(344, 224);
            this.buttonCancel.Name = "buttonCancel";
            this.buttonCancel.Size = new System.Drawing.Size(75, 23);
            this.buttonCancel.TabIndex = 8;
            this.buttonCancel.Text = "Cancel";
            this.buttonCancel.Click += new System.EventHandler(this.buttonCancel_Click);
            // 
            // buttonNext
            // 
            this.buttonNext.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonNext.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.buttonNext.Location = new System.Drawing.Point(260, 224);
            this.buttonNext.Name = "buttonNext";
            this.buttonNext.Size = new System.Drawing.Size(75, 23);
            this.buttonNext.TabIndex = 7;
            this.buttonNext.Text = "&Next >";
            this.buttonNext.Click += new System.EventHandler(this.buttonNext_Click);
            // 
            // buttonBack
            // 
            this.buttonBack.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonBack.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.buttonBack.Location = new System.Drawing.Point(184, 224);
            this.buttonBack.Name = "buttonBack";
            this.buttonBack.Size = new System.Drawing.Size(75, 23);
            this.buttonBack.TabIndex = 6;
            this.buttonBack.Text = "< &Back";
            this.buttonBack.Click += new System.EventHandler(this.buttonBack_Click);
            // 
            // buttonHelp
            // 
            this.buttonHelp.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.buttonHelp.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.buttonHelp.Location = new System.Drawing.Point(8, 224);
            this.buttonHelp.Name = "buttonHelp";
            this.buttonHelp.Size = new System.Drawing.Size(75, 23);
            this.buttonHelp.TabIndex = 9;
            this.buttonHelp.Text = "&Help";
            this.buttonHelp.Visible = false;
            this.buttonHelp.Click += new System.EventHandler(this.buttonHelp_Click);
            // 
            // Wizard
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Inherit;
            this.Controls.Add(this.buttonHelp);
            this.Controls.Add(this.buttonCancel);
            this.Controls.Add(this.buttonNext);
            this.Controls.Add(this.buttonBack);
            this.Name = "Wizard";
            this.Size = new System.Drawing.Size(428, 256);
            this.ResumeLayout(false);

        }
		#endregion

		#region Properties
		/// <summary>
		/// Gets or sets which edge of the parent container a control is docked to.
		/// </summary>
		[DefaultValue(DockStyle.Fill)]
		[Category("Layout")]
		[Description("Gets or sets which edge of the parent container a control is docked to.")]
		public new DockStyle Dock
		{
			get
			{
				return base.Dock;
			}
			set
			{
				base.Dock = value;
			}
		}

		/// <summary>
		/// Gets the collection of wizard pages in this tab control.
		/// </summary>
		[Category("Wizard")]
		[DesignerSerializationVisibility(DesignerSerializationVisibility.Content)]
		[Description("Gets the collection of wizard pages in this tab control.")]
		public WizardPagesCollection Pages
		{
			get
			{
				return _pages;
			}
		}

		/// <summary>
		/// Gets or sets the currently-selected wizard page.
		/// </summary>
		[Browsable(false)]
		[DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
		public WizardPage SelectedPage
		{
			get
			{
				return _selectedPage;
			}
			set
			{
				// select new page
				ActivatePage(value);
			}
		}

		/// <summary>
		/// Gets or sets the currently-selected wizard page by index.
		/// </summary>
		[Browsable(false)]
		[DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
		internal int SelectedIndex
		{
			get
			{
				return _pages.IndexOf(_selectedPage);
			}
			set
			{
				// check if there are any pages
				if(_pages.Count == 0)
				{
					// reset invalid index
					ActivatePage(-1);
					return;
				}

				// validate page index
				if (value < -1 || value >= _pages.Count)
				{
                    throw new IndexOutOfRangeException(
                        string.Format("The page index ({0}) is out of range. Should be between 0 and {1}", value, Convert.ToString(_pages.Count - 1)));
				}

				// select new page
				ActivatePage(value);
			}
		}

		/// <summary>
		/// Gets or sets the image displayed on the header of the standard pages.
		/// </summary>
		[DefaultValue(null)]
		[Category("Wizard")]
		[Description("Gets or sets the image displayed on the header of the standard pages.")]
		public Image HeaderImage
		{
			get
			{
				return _headerImage;
			}
			set
			{
				if (_headerImage != value)
				{
					_headerImage = value;
					Invalidate();
				}
			}
		}

		/// <summary>
		/// Gets or sets the image displayed on the welcome and finish pages.
		/// </summary>
		[DefaultValue(null)]
		[Category("Wizard")]
		[Description("Gets or sets the image displayed on the welcome and finish pages.")]
		public Image WelcomeImage
		{
			get
			{
				return _welcomeImage;
			}
			set
			{
				if (_welcomeImage != value)
				{
					_welcomeImage = value;
					Invalidate();
				}
			}
		}

		/// <summary>
		/// Gets or sets the font used to display the description of a standard page.
		/// </summary>
		[Category("Appearance")]
		[Description("Gets or sets the font used to display the description of a standard page.")]
		public Font HeaderFont
		{
			get
			{
			    return _headerFont ?? Font;
			}
		    set
			{
				if (!Equals(_headerFont, value))
				{
					_headerFont = value;
					Invalidate();
				}
			}
		}
		protected bool ShouldSerializeHeaderFont()
		{
			return _headerFont != null;
		}

		/// <summary>
		/// Gets or sets the font used to display the title of a standard page.
		/// </summary>
		[Category("Appearance")]
		[Description("Gets or sets the font used to display the title of a standard page.")]
		public Font HeaderTitleFont
		{
			get
			{
			    return _headerTitleFont ?? new Font(Font.FontFamily, Font.Size + 2, FontStyle.Bold);
			}
		    set
			{
				if (!Equals(_headerTitleFont, value))
				{
					_headerTitleFont = value;
					Invalidate();
				}
			}
		}
		protected bool ShouldSerializeHeaderTitleFont()
		{
			return _headerTitleFont != null;
		}

		/// <summary>
		/// Gets or sets the font used to display the description of a welcome of finish page.
		/// </summary>
		[Category("Appearance")]
		[Description("Gets or sets the font used to display the description of a welcome of finish page.")]
		public Font WelcomeFont
		{
			get
			{
			    return _welcomeFont ?? Font;
			}
		    set
			{
				if (!Equals(_welcomeFont, value))
				{
					_welcomeFont = value;
					Invalidate();
				}
			}
		}
		protected bool ShouldSerializeWelcomeFont()
		{
			return _welcomeFont != null;
		}

		/// <summary>
		/// Gets or sets the font used to display the title of a welcome of finish page.
		/// </summary>
		[Category("Appearance")]
		[Description("Gets or sets the font used to display the title of a welcome of finish page.")]
		public Font WelcomeTitleFont
		{
			get
			{
			    return _welcomeTitleFont ?? new Font(Font.FontFamily, Font.Size + 10, FontStyle.Bold);
			}
		    set
			{
				if (!Equals(_welcomeTitleFont, value))
				{
					_welcomeTitleFont = value;
					Invalidate();
				}
			}
		}
		protected bool ShouldSerializeWelcomeTitleFont()
		{
			return _welcomeTitleFont != null;
		}

		/// <summary>
		/// Gets or sets the enabled state of the Next button. 
		/// </summary>
		[Browsable(false)]
		[DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
		public bool NextEnabled
		{
			get
			{
				return buttonNext.Enabled;
			}
			set
			{
				buttonNext.Enabled = value;
			}
		}

		/// <summary>
		/// Gets or sets the enabled state of the back button. 
		/// </summary>
		[Browsable(false)]
		[DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
		public bool BackEnabled
		{
			get
			{
				return buttonBack.Enabled;
			}
			set
			{
				buttonBack.Enabled = value;
			}
		}

		/// <summary>
		/// Gets or sets the enabled state of the cancel button. 
		/// </summary>
		[Browsable(false)]
		[DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
		public bool CancelEnabled
		{
			get
			{
				return buttonCancel.Enabled;
			}
			set
			{
		
				buttonCancel.Enabled = value;
			}
		}

		/// <summary>
		/// Gets or sets the visible state of the help button. 
		/// </summary>
		[Category("Behavior")]
		[DefaultValue(false)]
		[Description("Gets or sets the visible state of the help button. ")]
		public bool HelpVisible
		{
			get
			{
				return buttonHelp.Visible;
			}
			set
			{
		
				buttonHelp.Visible = value;
			}
		}

		/// <summary>
		/// Gets or sets the text displayed by the Next button. 
		/// </summary>
		[Browsable(false)]
		[DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
		public string NextText
		{
			get
			{
				return buttonNext.Text;
			}
			set
			{
				buttonNext.Text = value;
			}
		}

		/// <summary>
		/// Gets or sets the text displayed by the back button. 
		/// </summary>
		[Browsable(false)]
		[DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
		public string BackText
		{
			get
			{
				return buttonBack.Text;
			}
			set
			{
				buttonBack.Text = value;
			}
		}

		/// <summary>
		/// Gets or sets the text displayed by the cancel button. 
		/// </summary>
		[Browsable(false)]
		[DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
		public string CancelText
		{
			get
			{
				return buttonCancel.Text;
			}
			set
			{
		
				buttonCancel.Text = value;
			}
		}

		/// <summary>
		/// Gets or sets the text displayed by the cancel button. 
		/// </summary>
		[Browsable(false)]
		[DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
		public string HelpText
		{
			get
			{
				return buttonHelp.Text;
			}
			set
			{
		
				buttonHelp.Text = value;
			}
		}

		#endregion

		#region Methods
		/// <summary>
		/// Swithes forward to next wizard page.
		/// </summary>
		public void Next()
		{
		    if (DesignMode && SelectedIndex == _pages.Count - 1)
		    {
		        buttonNext.Enabled = false;
		        return;
		    }

            // check if button is finish mode
		    if (buttonNext.DialogResult == DialogResult.OK)
		    {
		        OnFinish(new CancelEventArgs());
		        return;
		    }

            // handle page switch
            OnBeforeSwitchPages(new BeforeSwitchPagesEventArgs { OldIndex = SelectedIndex, NewIndex = SelectedIndex + 1 });
		}

		/// <summary>
		/// Swithes backward to previous wizard page.
		/// </summary>
		public void Back()
		{
			if (SelectedIndex == 0)
			{
				buttonBack.Enabled = false;
			    return;
			}

			// handle page switch
			OnBeforeSwitchPages(new BeforeSwitchPagesEventArgs { OldIndex = SelectedIndex, NewIndex = SelectedIndex - 1});
		}

		/// <summary>
		/// Activates the specified wizard bage.
		/// </summary>
		/// <param name="index">An Integer value representing the zero-based index of the page to be activated.</param>
		private void ActivatePage(int index)
		{
			// check if new page is invalid
			if (index < 0 || index >= _pages.Count)
			{
				// filter out
				return;
			}
		
			// get new page
			var page = _pages[index];

			// activate page
			ActivatePage(page);
		}

		/// <summary>
		/// Activates the specified wizard bage.
		/// </summary>
		/// <param name="page">A WizardPage object representing the page to be activated.</param>
		private void ActivatePage(WizardPage page)
		{
			// validate given page
			if (_pages.Contains(page) == false)
			{
				// filter out
				return;
			}

			// deactivate current page
			if (_selectedPage != null)
			{
				_selectedPage.Visible = false;
			}

			// activate new page
			_selectedPage = page;

			if (_selectedPage != null)
			{
				//Ensure that this panel displays inside the wizard
				_selectedPage.Parent = this;
				if (Contains(_selectedPage) == false)
				{	
					Container.Add(_selectedPage);
				}

				if (SelectedIndex < _pages.Count - 1)
				{
                    buttonNext.Text = "&Next >";
                    buttonNext.DialogResult = DialogResult.None;

				    if (DesignMode)
				        buttonNext.Enabled = true;
				}
				else
				{
                    buttonNext.Text = "Finish";
                    buttonNext.DialogResult = DialogResult.OK;

                    if (DesignMode)
                        buttonNext.Enabled = false;
                }

				//Make it fill the space
				_selectedPage.SetBounds(0, 0, Width, Height - FooterAreaHeight);
				_selectedPage.Visible = true;
				_selectedPage.BringToFront();
				FocusFirstTabIndex(_selectedPage);
			}
			
			//What should the back button say
			buttonBack.Enabled = SelectedIndex > 0;

			// refresh
			if (_selectedPage != null)
			{
				_selectedPage.Invalidate();
			}
			else
			{
				Invalidate();
			}
		}

		/// <summary>
		/// Focus the control with a lowest tab index in the given container.
		/// </summary>
		/// <param name="container">A Control object to pe processed.</param>
		private void FocusFirstTabIndex(Control container)
		{
			// init search result varialble
			Control searchResult = null;

			// find the control with the lowest tab index
			foreach (Control control in container.Controls)
			{
				if (control.CanFocus && (searchResult == null || control.TabIndex < searchResult.TabIndex))
				{
					searchResult = control;
				}
			}

			// check if anything searchResult
			if (searchResult != null)
			{
				// focus found control
				searchResult.Focus();
			}
			else
			{
				// focus the container
				container.Focus();
			}
		}

		/// <summary>
		/// Raises the SwitchPages event.
		/// </summary>
		/// <param name="e">A WizardPageEventArgs object that holds event data.</param>
		protected virtual void OnBeforeSwitchPages(BeforeSwitchPagesEventArgs e)
		{
			// check if there are subscribers
			if (BeforeSwitchPages != null)
			{
				// raise BeforeSwitchPages event
				BeforeSwitchPages(this, e);
			}

			// check if user canceled
			if (e.Cancel)
			{
				// filter
				return;
			}

			// activate new page
			ActivatePage(e.NewIndex);

			// raise the after event
			OnAfterSwitchPages(new AfterSwitchPagesEventArgs { OldIndex =  e.OldIndex, NewIndex = e.NewIndex});
		}

		/// <summary>
		/// Raises the SwitchPages event.
		/// </summary>
		/// <param name="e">A WizardPageEventArgs object that holds event data.</param>
		protected virtual void OnAfterSwitchPages(AfterSwitchPagesEventArgs e)
		{
			// check if there are subscribers
			if (AfterSwitchPages != null)
			{
				// raise AfterSwitchPages event
				AfterSwitchPages(this, e);
			}
		}

		/// <summary>
		/// Raises the Cancel event.
		/// </summary>
		/// <param name="e">A CancelEventArgs object that holds event data.</param>
		protected virtual void OnCancel(CancelEventArgs e)
		{
			// check if there are subscribers
			if (Cancel != null)
			{
				// raise Cancel event
				Cancel(this, e);
			}

			// check if user canceled
			if (e.Cancel)
			{
				// cancel closing (when ShowDialog is used)
				ParentForm.DialogResult = DialogResult.None;
			}
			else
			{
				// ensure parent form is closed (even when ShowDialog is not used)
				ParentForm.Close();
			}
		}

		/// <summary>
		/// Raises the Finish event.
		/// </summary>
		/// <param name="e">A EventArgs object that holds event data.</param>
		protected virtual void OnFinish(CancelEventArgs e)
		{
			// check if there are subscribers
			if (Finish != null)
			{
				// raise Finish event
				Finish(this, e);
			}

            // check if user canceled
            if (e.Cancel)
            {
                // cancel closing (when ShowDialog is used)
                ParentForm.DialogResult = DialogResult.None;
            }
            else
            {
                // ensure parent form is closed (even when ShowDialog is not used)
                ParentForm.Close();
            }
		}

		/// <summary>
		/// Raises the Help event.
		/// </summary>
		/// <param name="e">A EventArgs object that holds event data.</param>
		protected virtual void OnHelp(EventArgs e)
		{
			// check if there are subscribers
			if (Help != null)
			{
				// raise Help event
				Help(this, e);
			}
		}

		/// <summary>
		/// Raises the Load event.
		/// </summary>
		protected override void OnLoad(EventArgs e)
		{
			// raise the Load event
			base.OnLoad(e);

			Graphics g = this.CreateGraphics();
			float DpiXfactor = g.DpiX / 96;
			float DpiYfactor = g.DpiY / 96;
			FooterAreaHeight = (int)(48 * DpiYfactor);
			_offsetCancel = new Point((int)(100 * DpiXfactor), (int)(36 * DpiYfactor));
			_offsetNext = new Point((int)((100 + 1*75 + 20) * DpiXfactor), (int)(36 * DpiYfactor));
			_offsetBack = new Point((int)((100 + 2*75 + 30) * DpiXfactor), (int)(36 * DpiYfactor));

			// activate first page, if exists
			if (_pages.Count > 0)
			{
				ActivatePage(0);
			}
		}

		/// <summary>
		/// Raises the Resize event.
		/// </summary>
		protected override void OnResize(EventArgs e)
		{
			// raise the Resize event
			base.OnResize(e);
		}

		/// <summary>
		/// Raises the Paint event.
		/// </summary>
		protected override void OnPaint(PaintEventArgs e)
		{
			// resize the selected page to fit the wizard
			if (_selectedPage != null)
			{
				_selectedPage.SetBounds(0, 0, Width, Height - FooterAreaHeight);
			}

			// position navigation buttons
			buttonCancel.Location = new Point(Width - _offsetCancel.X,
				Height - _offsetCancel.Y);
			buttonNext.Location = new Point(Width - _offsetNext.X,
				Height - _offsetNext.Y);
			buttonBack.Location = new Point(Width - _offsetBack.X,
				Height - _offsetBack.Y);
			buttonHelp.Location = new Point(buttonHelp.Location.X,
				Height - _offsetBack.Y);

			// raise the Paint event
			base.OnPaint(e);
			
			Rectangle bottomRect = ClientRectangle;
			bottomRect.Y = Height - FooterAreaHeight;
			bottomRect.Height = FooterAreaHeight;
			ControlPaint.DrawBorder3D(e.Graphics, bottomRect, Border3DStyle.Etched, Border3DSide.Top);
		}

		/// <summary>
		/// Raises the ControlAdded event.
		/// </summary>
		protected override void OnControlAdded(ControlEventArgs e) 
		{
			// prevent other controls from being added directly to the wizard
			if (e.Control is WizardPage == false &&
				e.Control != buttonCancel &&
				e.Control != buttonNext &&
				e.Control != buttonBack)
			{
				// add the control to the selected page
				if (_selectedPage != null)
				{
					_selectedPage.Controls.Add(e.Control);
				}
			}
			else
			{
				// raise the ControlAdded event
				base.OnControlAdded(e);
			}
		}

		#endregion

		#region Events
		/// <summary>
		/// Occurs before the wizard pages are switched, giving the user a chance to validate.
		/// </summary>
		[Category("Wizard")]
		[Description("Occurs before the wizard pages are switched, giving the user a chance to validate.")]
		public event BeforeSwitchPagesEventHandler BeforeSwitchPages;
		/// <summary>
		/// Occurs after the wizard pages are switched, giving the user a chance to setup the new page.
		/// </summary>
		[Category("Wizard")]
		[Description("Occurs after the wizard pages are switched, giving the user a chance to setup the new page.")]
		public event AfterSwitchPagesEventHandler AfterSwitchPages;
		/// <summary>
		/// Occurs when wizard is canceled, giving the user a chance to validate.
		/// </summary>
		[Category("Wizard")]
		[Description("Occurs when wizard is canceled, giving the user a chance to validate.")]
		public event CancelEventHandler Cancel;
		/// <summary>
		/// Occurs when wizard is finished, giving the user a chance to do extra stuff.
		/// </summary>
		[Category("Wizard")]
		[Description("Occurs when wizard is finished, giving the user a chance to do extra stuff.")]
		public event CancelEventHandler Finish;
		/// <summary>
		/// Occurs when the user clicks the help button.
		/// </summary>
		[Category("Wizard")]
		[Description("Occurs when the user clicks the help button.")]
		public event EventHandler Help;
		/// <summary>
		/// Represents the method that will handle the BeforeSwitchPages event of the Wizard control.
		/// </summary>
		public delegate void BeforeSwitchPagesEventHandler(object sender, BeforeSwitchPagesEventArgs e);
		/// <summary>
		/// Represents the method that will handle the AfterSwitchPages event of the Wizard control.
		/// </summary>
		public delegate void AfterSwitchPagesEventHandler(object sender, AfterSwitchPagesEventArgs e);
		#endregion

		#region Events handlers
		/// <summary>
		/// Handles the Click event of buttonNext.
		/// </summary>
		private void buttonNext_Click(object sender, EventArgs e)
		{
			Next();
		}

		/// <summary>
		/// Handles the Click event of buttonBack.
		/// </summary>
		private void buttonBack_Click(object sender, EventArgs e)
		{
			Back();
		}
		
		/// <summary>
		/// Handles the Click event of buttonCancel.
		/// </summary>
		private void buttonCancel_Click(object sender, EventArgs e)
		{
			// check if button is cancel mode
			OnCancel(new CancelEventArgs());
		}
		
		/// <summary>
		/// Handles the Click event of buttonHelp.
		/// </summary>
        private void buttonHelp_Click(object sender, EventArgs e)
        {
			OnHelp(EventArgs.Empty);
        }
		#endregion

		#region Inner classes
		/// <summary>
		/// Represents a designer for the wizard control.
		/// </summary>
		internal class WizardDesigner : ParentControlDesigner
		{

			#region Methods
			/// <summary>
			/// Overrides the handling of Mouse clicks to allow back-next to work in the designer.
			/// </summary>
			/// <param name="msg">A Message value.</param>
			protected override void WndProc(ref Message msg)
			{
				// declare PInvoke constants
				const int WM_LBUTTONDOWN = 0x0201;
				const int WM_LBUTTONDBLCLK = 0x0203;

				// check message
				if (msg.Msg == WM_LBUTTONDOWN || msg.Msg == WM_LBUTTONDBLCLK)
				{
					// get the control under the mouse
					var ss = (ISelectionService)GetService(typeof(ISelectionService));

				    var wizard = ss.PrimarySelection as Wizard;
				    if (wizard != null)
					{
						// extract the mouse position
						int xPos = (short)((uint)msg.LParam & 0x0000FFFF);
						int yPos = (short)(((uint)msg.LParam & 0xFFFF0000) >> 16);
						Point mousePos =  new Point(xPos, yPos);
						
						if (msg.HWnd == wizard.buttonNext.Handle)
						{
							if (wizard.buttonNext.Enabled && 
								wizard.buttonNext.ClientRectangle.Contains(mousePos))
							{
								//Press the button
								wizard.Next();
							}
						}
						else if (msg.HWnd == wizard.buttonBack.Handle)
						{
							if (wizard.buttonBack.Enabled && 
								wizard.buttonBack.ClientRectangle.Contains(mousePos))
							{
								//Press the button
								wizard.Back();
							}
						}
						
						// filter message
						return;
					}
				}

				// forward message
				base.WndProc(ref msg);
			}

			/// <summary>
			/// Prevents the grid from being drawn on the Wizard.
			/// </summary>
			protected override bool DrawGrid
			{
				get
				{
					return false;
				}
			}
			#endregion

		}

		/// <summary>
		/// Provides data for the AfterSwitchPages event of the Wizard control.
		/// </summary>
		public class AfterSwitchPagesEventArgs : EventArgs
		{
			/// <summary>
			/// Gets the index of the old page.
			/// </summary>
			public int OldIndex { get; set; }

		    /// <summary>
			/// Gets or sets the index of the new page.
			/// </summary>
			public int NewIndex { get; set; }
		}

		/// <summary>
		/// Provides data for the BeforeSwitchPages event of the Wizard control.
		/// </summary>
        public class BeforeSwitchPagesEventArgs : EventArgs
		{
			/// <summary>
			/// Indicates whether the page switch should be canceled.
			/// </summary>
			public bool Cancel { get; set; }

            /// <summary>
            /// Gets the index of the old page.
            /// </summary>
            public int OldIndex { get; set; }

            /// <summary>
			/// Gets or sets the index of the new page.
			/// </summary>
			public int NewIndex { get; set; }


		}
		#endregion

	}
}
