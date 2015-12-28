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
                writer.WriteStartElement("visio", "PublishAddinVSTO", "http://schemas.microsoft.com/wix/Visio");

                if (Id != null) 
                    writer.WriteAttributeString("Id", Id);

                if (FriendlyName != null)
                    writer.WriteAttributeString("FriendlyName", FriendlyName);

                if (Description != null)
                    writer.WriteAttributeString("Description", Description);

                writer.WriteAttributeString("CommandLineSafe", CommandLineSafe != 0 ? "yes" : "no");

                writer.WriteAttributeString("LoadBehavior", LoadBehavior.ToString());

                if (VisioEditionAttribute == (int)VisioEdition.X86)
                    writer.WriteAttributeString("VisioEdition", "x86");

                if (VisioEditionAttribute == (int)VisioEdition.X64)
                    writer.WriteAttributeString("VisioEdition", "x64");

                writer.WriteEndElement();
            }

            public WixSerialize.ISchemaElement ParentElement { get; set; }

            public string Id { get; set; }
            public string FriendlyName { get; set; }
            public string Description { get; set; }
            public int CommandLineSafe { get; set; }
            public int LoadBehavior { get; set; }
            public int VisioEditionAttribute { get; set; }
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
                    Id = row[0].ToString(),
                    FriendlyName = row[2].ToString(),
                    Description = row[3].ToString(),
                    VisioEditionAttribute = int.Parse(row[4].ToString()),
                    CommandLineSafe = int.Parse(row[5].ToString()),
                    LoadBehavior = int.Parse(row[6].ToString())
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
