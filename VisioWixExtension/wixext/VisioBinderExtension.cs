using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using Extensibility;
using Microsoft.Tools.WindowsInstallerXml;

namespace VisioWixExtension
{
    /// <summary>
    /// The decompiler for the Windows Installer XML Toolset Gaming Extension.
    /// </summary>
    public sealed class VisioBinderExtension : BinderExtension
    {
        public override void DatabaseFinalize(Output output)
        {
            HarvestAddinRegistrationsFromSourceFiles(output);

            base.DatabaseFinalize(output);
        }

        private void SetDefaultFriendlyNameFromAssemblyTitle(Row row, Assembly assembly)
        {
            var addinFriendlyName = row[(int)TableFields.arqFriendlyName].ToString();

            if (!string.IsNullOrEmpty(addinFriendlyName))
                return;

            var assemblyTitle = Attribute.GetCustomAttribute(assembly, typeof (AssemblyTitleAttribute), false)
                as AssemblyTitleAttribute;

            if (assemblyTitle != null)
                row[(int)TableFields.arqFriendlyName] = assemblyTitle.Title;
        }

        private void SetDefaultDescriptionFromAssemblyDescription(Row row, Assembly assembly)
        {
            var addinDescription = row[(int)TableFields.arqDescription].ToString();

            if (!string.IsNullOrEmpty(addinDescription))
                return;

            var assemblyDescription = Attribute.GetCustomAttribute(assembly, typeof (AssemblyDescriptionAttribute), false) 
                as AssemblyDescriptionAttribute;

            if (assemblyDescription != null)
                row[(int)TableFields.arqDescription] = assemblyDescription.Description;
        }
        
        private void HarvestAddinRegistrationsFromSourceFiles(Output output)
        {
            var fileTable = output.Tables["WixFile"];
            if (fileTable == null)
                return;

            var addinRegistrationTable = output.Tables["AddinRegistration"];
            if (addinRegistrationTable == null)
                return;

            var paths = fileTable.Rows
                .Cast<WixFileRow>()
                .ToDictionary(item => item.File, item => item.Source);

            foreach (Row row in addinRegistrationTable.Rows)
            {
                var fileId = row[(int) TableFields.arqFile].ToString();

                string filePath;
                if (!paths.TryGetValue(fileId, out filePath))
                {
                    Core.OnMessage(VisioErrors.FileIdentifierNotFound(fileId));
                    continue;
                }

                if (Path.GetExtension(filePath) == ".vsto")
                {
                    var addinFilePath = Path.ChangeExtension(filePath, ".dll");

                    if (!File.Exists(addinFilePath))
                        continue;

                    var assembly = Assembly.LoadFrom(addinFilePath);

                    SetDefaultFriendlyNameFromAssemblyTitle(row, assembly);
                    SetDefaultDescriptionFromAssemblyDescription(row, assembly);
                }
                else
                {
                    var assembly = Assembly.LoadFrom(filePath);

                    var addin = assembly
                        .GetTypes()
                        .FirstOrDefault(t => t.IsClass && typeof(IDTExtensibility2).IsAssignableFrom(t));

                    if (addin == null)
                    {
                        Core.OnMessage(VisioErrors.AddinNotFound(filePath));
                        continue;
                    }

                    SetDefaultFriendlyNameFromAssemblyTitle(row, assembly);
                    SetDefaultDescriptionFromAssemblyDescription(row, assembly);

                    row[(int)TableFields.arqProgId] = Marshal.GenerateProgIdForType(addin);
                    row[(int)TableFields.arqClassId] = "{" + Marshal.GenerateGuidForType(addin).ToString().ToUpper(CultureInfo.InvariantCulture) + "}";
                    row[(int)TableFields.arqClass] = addin.FullName;
                    row[(int)TableFields.arqAssembly] = addin.Assembly.FullName;
                    row[(int)TableFields.arqVersion] = addin.Assembly.GetName().Version.ToString();
                    row[(int)TableFields.arqRuntimeVersion] = addin.Assembly.ImageRuntimeVersion;
                }
            }
        }
    }
}
