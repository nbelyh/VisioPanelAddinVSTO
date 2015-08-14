//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated by a tool.
//     Runtime Version:2.0.50727.5485
//
//     Changes to this file may cause incorrect behavior and will be lost if
//     the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace VisioWixExtension
{
    using System;
    using System.Reflection;
    using System.Resources;
    using Microsoft.Tools.WindowsInstallerXml;
    
    
    public class VisioErrorEventArgs : WixErrorEventArgs
    {
        
        private static ResourceManager resourceManager = new ResourceManager("VisioWixExtension.Data.Messages", Assembly.GetExecutingAssembly());
        
        public VisioErrorEventArgs(SourceLineNumberCollection sourceLineNumbers, int id, string resourceName, params object[] messageArgs) : 
                base(sourceLineNumbers, id, resourceName, messageArgs)
        {
        }
        
        public override ResourceManager ResourceManager
        {
            get
            {
                return resourceManager;
            }
        }
    }
    
    public sealed class VisioErrors
    {
        
        private VisioErrors()
        {
        }
        
        public static VisioErrorEventArgs InvalidVisioEdition(SourceLineNumberCollection sourceLineNumbers, string value)
        {
            return new VisioErrorEventArgs(sourceLineNumbers, 6501, "VisioErrors_InvalidVisioEdition_1", value);
        }
        
        public static VisioErrorEventArgs InvalidLanguage(SourceLineNumberCollection sourceLineNumbers, string value)
        {
            return new VisioErrorEventArgs(sourceLineNumbers, 6502, "VisioErrors_InvalidLanguage_1", value);
        }
        
        public static VisioErrorEventArgs InvalidUint(SourceLineNumberCollection sourceLineNumbers, string attributeName, string value)
        {
            return new VisioErrorEventArgs(sourceLineNumbers, 6503, "VisioErrors_InvalidUint_1", attributeName, value);
        }
        
        public static VisioErrorEventArgs UsingPublish(SourceLineNumberCollection sourceLineNumbers)
        {
            return new VisioErrorEventArgs(sourceLineNumbers, 6508, "VisioErrors_UsingPublish_1");
        }
    }
    
    public class VisioWarningEventArgs : WixWarningEventArgs
    {
        
        private static ResourceManager resourceManager = new ResourceManager("VisioWixExtension.Data.Messages", Assembly.GetExecutingAssembly());
        
        public VisioWarningEventArgs(SourceLineNumberCollection sourceLineNumbers, int id, string resourceName, params object[] messageArgs) : 
                base(sourceLineNumbers, id, resourceName, messageArgs)
        {
        }
        
        public override ResourceManager ResourceManager
        {
            get
            {
                return resourceManager;
            }
        }
    }
    
    public sealed class VisioWarnings
    {
        
        private VisioWarnings()
        {
        }
        
        public static VisioWarningEventArgs InvalidFileExtension(SourceLineNumberCollection sourceLineNumbers, string fileName, string expectedVisioContentType)
        {
            return new VisioWarningEventArgs(sourceLineNumbers, 6507, "VisioWarnings_InvalidFileExtension_1", fileName, expectedVisioContentType);
        }
    }
}