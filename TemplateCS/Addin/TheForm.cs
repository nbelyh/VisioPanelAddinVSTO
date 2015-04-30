using System.Windows.Forms;
using Visio = Microsoft.Office.Interop.Visio;

namespace $csprojectname$
{
    public partial class TheForm : Form
    {
        private readonly Visio.Window _window;

        /// <summary>
        /// Form constructor, receives parent Visio diagram window
        /// </summary>
        /// <param name="window">Parent Visio diagram window</param>
        public TheForm(Visio.Window window)
        {
            _window = window;
            InitializeComponent();
        }

        /// <summary>
        /// Sample method. We just show a Message Box. 
        /// Do something meaningful here instead.
        /// </summary>
        private void button1_Click(object sender, System.EventArgs e)
        {
            MessageBox.Show(_window.Document.Name);
        }
    }
}
