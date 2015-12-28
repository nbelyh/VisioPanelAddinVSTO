using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PanelAddinWizard
{
    public class WixSetupOptions
    {
        public bool EnableWixSetup { get; set; }
        public bool AddVisioFiles { get; set; }
        public bool CreateNewVisioFiles { get; set; }
        public bool DuplicateExistingVisioFiles { get; set; }
        public string[] VisioFilePaths { get; set; }

        public bool EnableWixUI { get; set; }
        public string WixUI { get; set; }

        public bool HaveVisioFiles { get { return EnableWixSetup && !CreateNewVisioFiles && VisioFilePaths != null && VisioFilePaths.Length > 0; } }
    }

}
