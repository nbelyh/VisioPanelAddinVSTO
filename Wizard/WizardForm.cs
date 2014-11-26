using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Windows.Forms;

namespace PanelAddinWizard
{
    public partial class WizardForm : Form
    {
        public WizardForm()
        {
            InitializeComponent();
            checkSupportTaskPane.DataBindings.Add("Checked", this, "TaskPane");
            checkSupportRibbon.DataBindings.Add("Checked", this, "Ribbon");
            checkSupportCommandBars.DataBindings.Add("Checked", this, "CommandBars");
        }

        public bool TaskPane { get; set; }
        public bool Ribbon { get; set; }
        public bool CommandBars { get; set; }
    }
}
