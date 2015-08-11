using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PanelAddinWizard
{
    public class WixSetupOptions
    {
        public bool Enabled { get; set; }
        public bool CreateNew { get; set; }
        public bool Duplicate { get; set; }
        public string[] Paths { get; set; }
    }

}
