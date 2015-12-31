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
            try
            {
                HarvestAddinRegistrationsFromSourceFiles(output);
            }
            catch (Exception e)
            {
                Core.OnMessage(VisioErrors.InternalException(e.ToString()));
            }

            base.DatabaseFinalize(output);
        }

        private void SetDefaultFriendlyNameFromAssemblyTitle(Row row, Lazy<Assembly> assembly)
        {
            var addinFriendlyName = (string) row[(int)TableFields.arqFriendlyName];

            if (!string.IsNullOrEmpty(addinFriendlyName))
                return;

            var assemblyTitle = Attribute.GetCustomAttribute(assembly.Value, typeof (AssemblyTitleAttribute), false)
                as AssemblyTitleAttribute;

            if (assemblyTitle != null)
                row[(int)TableFields.arqFriendlyName] = assemblyTitle.Title;
        }

        private void SetDefaultDescriptionFromAssemblyDescription(Row row, Lazy<Assembly> assembly)
        {
            var addinDescription = (string) row[(int)TableFields.arqDescription];

            if (!string.IsNullOrEmpty(addinDescription))
                return;

            var assemblyDescription = Attribute.GetCustomAttribute(assembly.Value, typeof (AssemblyDescriptionAttribute), false) 
                as AssemblyDescriptionAttribute;

            if (assemblyDescription != null)
                row[(int)TableFields.arqDescription] = assemblyDescription.Description;
        }
        
        AddinType GetAddinTypeFromFilePath(string filePath)
        {
            var fileExtension = Path.GetExtension(filePath);
            if (fileExtension == null)
                return AddinType.Unknown;

            switch (fileExtension.ToLower(CultureInfo.InvariantCulture))
            {
                case ".vsto":
                    return AddinType.VSTO;

                case ".dll":
                    return AddinType.COM;
                    
                default:
                    return AddinType.Unknown;
            }
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
                var fileId = (string) row[(int) TableFields.arqFile];
                var addinType = (AddinType) (int) row[(int) TableFields.arqAddinType];

                string filePath;
                if (!paths.TryGetValue(fileId, out filePath))
                {
                    Core.OnMessage(VisioErrors.FileIdentifierNotFound(fileId));
                    continue;
                }

                if (addinType == AddinType.Unknown)
                {
                    var detectedAddinType = GetAddinTypeFromFilePath(filePath);
                    if (detectedAddinType == AddinType.Unknown)
                    {
                        Core.OnMessage(VisioErrors.UnknownAddinType(fileId));
                        continue;
                    }

                    addinType = detectedAddinType;
                    row[(int) TableFields.arqAddinType] = (int) addinType;
                }

                if (addinType == AddinType.VSTO)
                {
                    var addinFilePath = Path.ChangeExtension(filePath, ".dll");

                    if (!File.Exists(addinFilePath))
                        continue;

                    var assembly = new Lazy<Assembly>( () => Assembly.Load(File.ReadAllBytes(addinFilePath)));

                    SetDefaultFriendlyNameFromAssemblyTitle(row, assembly);
                    SetDefaultDescriptionFromAssemblyDescription(row, assembly);
                }

                if (addinType == AddinType.COM)
                {
                    var assembly = new Lazy<Assembly>(() => Assembly.Load(File.ReadAllBytes(filePath)));

                    var addin = assembly.Value
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
                    row[(int)TableFields.arqAssembly] = assembly.Value.FullName;
                    row[(int)TableFields.arqVersion] = assembly.Value.GetName().Version.ToString();
                    row[(int)TableFields.arqRuntimeVersion] = assembly.Value.ImageRuntimeVersion;
                }
            }
        }
    }
}
