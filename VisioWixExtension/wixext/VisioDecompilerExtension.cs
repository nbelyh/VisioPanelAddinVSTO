using System.Resources;
using System.Xml;
using Microsoft.Tools.WindowsInstallerXml;
using WixSerialize = Microsoft.Tools.WindowsInstallerXml.Serialize;

namespace VisioWixExtension
{
    /// <summary>
    /// The decompiler for the Windows Installer XML Toolset Gaming Extension.
    /// </summary>
    public sealed class VisioDecompilerExtension : DecompilerExtension
    {
        /// <summary>
        /// Decompiles an extension table.
        /// </summary>
        /// <param name="table">The table to decompile.</param>
        public override void DecompileTable(Table table)
        {
            switch (table.Name)
            {
                case "AddinRegistration":
                    DecompileAddinRegistrationTable(table);
                    break;
                default:
                    base.DecompileTable(table);
                    break;
            }
        }

        class AddinRegistration : WixSerialize.ISchemaElement
        {
            public void OutputXml(XmlWriter writer)
            {
                writer.WriteStartElement("visio", "PublishAddin", "http://schemas.microsoft.com/wix/Visio");

                if (ProgId != null) 
                    writer.WriteAttributeString("ProgId", ProgId);

                if (FriendlyName != null)
                    writer.WriteAttributeString("FriendlyName", FriendlyName);

                if (Description != null)
                    writer.WriteAttributeString("Description", Description);

                writer.WriteAttributeString("CommandLineSafe", CommandLineSafe != 0 ? "yes" : "no");

                writer.WriteAttributeString("LoadBehavior", LoadBehavior.ToString());

                switch ((VisioEdition)VisioEditionAttribute)
                {
                    case VisioEdition.X86:
                        writer.WriteAttributeString("VisioEdition", "x86");
                        break;
                        
                    case VisioEdition.X64:
                        writer.WriteAttributeString("VisioEdition", "x64");
                        break;
                }

                switch ((AddinType)Type)
                {
                    case AddinType.COM:
                        writer.WriteAttributeString("Type", "COM");
                        break;

                    case AddinType.VSTO:
                        writer.WriteAttributeString("Type", "VSTO");
                        break;
                }

                writer.WriteEndElement();
            }

            public WixSerialize.ISchemaElement ParentElement { get; set; }

            public string ProgId { get; set; }
            public string FriendlyName { get; set; }
            public string Description { get; set; }
            public int CommandLineSafe { get; set; }
            public int LoadBehavior { get; set; }
            public int VisioEditionAttribute { get; set; }
            public int Type { get; set; }
        }

        /// <summary>
        /// Decompile the AddinRegistration table.
        /// </summary>
        /// <param name="table">The table to decompile.</param>
        private void DecompileAddinRegistrationTable(Table table)
        {
            foreach (Row row in table.Rows)
            {
                var file = (WixSerialize.File)Core.GetIndexedElement("File", (string)row[1]);

                var item = new AddinRegistration
                {
                    ProgId = row[0].ToString(),
                    FriendlyName = row[2].ToString(),
                    Description = row[3].ToString(),
                    VisioEditionAttribute = int.Parse(row[4].ToString()),
                    CommandLineSafe = int.Parse(row[5].ToString()),
                    LoadBehavior = int.Parse(row[6].ToString()),
                    Type = int.Parse(row[7].ToString())
                };

                if (null != file)
                {
                    file.AddChild(item);
                }
                else
                {
                    Core.OnMessage(WixWarnings.ExpectedForeignRow(row.SourceLineNumbers, table.Name, row.GetPrimaryKey(DecompilerCore.PrimaryKeyDelimiter), "File_", (string)row[1], "File"));
                }
            }
        }
    }
}
