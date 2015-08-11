using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using PanelAddinWizard;
using System.Xml;
using System.IO;

namespace PanelAddinWizardTestApp
{

    class Program
    {
        //static string GetFiles(string visioElement, string[] fileNames)
        //{
        //    if (fileNames == null)
        //        return "";

        //    var doc = new XmlDocument();

        //    const string PublishItemName = "Publish6EACFB1ABA5A4581A2F0DFA55A8B3445";

        //    var nodeRoot = doc.CreateElement("Root");
        //    foreach (var fileName in fileNames)
        //    {
        //        var nodeComponent = doc.CreateElement("Component");
        //        nodeRoot.AppendChild(nodeComponent);

        //        var nodeFile = doc.CreateElement("File");
        //        nodeFile.SetAttribute("Name", Path.GetFileName(fileName));
        //        nodeComponent.AppendChild(nodeFile);

        //        var nodePublish = doc.CreateElement(PublishItemName);
        //        nodePublish.SetAttribute("MenuPath",
        //            string.Format("$csprojectname$\\{0}", Path.GetFileNameWithoutExtension(fileName)));
        //        nodeFile.AppendChild(nodePublish);
        //    }

        //    return nodeRoot.InnerXml.Replace(PublishItemName, visioElement);
        //}

        [STAThread]
        static void Main(string[] args)
        {
            var wizardForm = new WizardForm(true)
            {
            };

            wizardForm.ShowDialog();

            wizardForm.GetSetupOptions();
            
            //var f = GetFiles("visio:PublishStencil", wizardForm.StencilPaths);
            //var g = GetFiles("visio:PublishTemplate", wizardForm.TemplatePaths);
        }
    }
}
