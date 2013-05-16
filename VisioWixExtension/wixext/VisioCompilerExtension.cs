using System;
using System.Collections.Generic;
using System.Reflection;
using System.Xml;
using System.Xml.Schema;
using Microsoft.Tools.WindowsInstallerXml;
using System.IO;
using System.Globalization;
using System.Text;

namespace VisioWixExtension
{
    internal class VisioCompilerExtension : CompilerExtension
    {
        private XmlSchema _schema;

        /// <summary>
        /// Gets the schema for this extension.
        /// </summary>
        /// <value>Schema for this extension.</value>
        public override XmlSchema Schema
        {
            get
            {
                if (null == _schema)
                    _schema = LoadXmlSchemaHelper(Assembly.GetExecutingAssembly(), "VisioWixExtension.Xsd.VisioWixExtension.xsd");

                return _schema;
            }
        }

        /// <summary>
        /// Processes an element for the Compiler.
        /// </summary>
        /// <param name="sourceLineNumbers">Source line number for the parent element.</param>
        /// <param name="parentElement">Parent element of element to process.</param>
        /// <param name="element">Element to process.</param>
        /// <param name="contextValues">Extra information about the context in which this element is being parsed.</param>
        public override void ParseElement(SourceLineNumberCollection sourceLineNumbers, XmlElement parentElement,
                                          XmlElement element, params string[] contextValues)
        {
            switch (parentElement.LocalName)
            {
                case "File":
                    string componentId = contextValues[1];

                    switch (element.LocalName)
                    {
                        case "PublishStencil":
                            ParseVisioElement(parentElement, element, componentId, VisioContentType.Stencil);
                            break;

                        case "PublishTemplate":
                            ParseVisioElement(parentElement, element, componentId, VisioContentType.Template);
                            break;

                        case "PublishHelp":
                            ParseVisioElement(parentElement, element, componentId, VisioContentType.Help);
                            break;

                        case "PublishAddon":
                            ParseVisioElement(parentElement, element, componentId, VisioContentType.Addon);
                            break;

                        default:
                            Core.UnexpectedElement(parentElement, element);
                            break;
                    }
                    break;
                default:
                    Core.UnexpectedElement(parentElement, element);
                    break;
            }
        }

        /// <summary>
        /// Supported Visio versions (2003/2007/2010/etc)
        /// </summary>
        [Flags]
        private enum VisioVersion
        {
            Visio2003 = 1,
            Visio2007 = 2,
            Visio2010 = 4,
            Visio2013 = 8,
        };

        /// <summary>
        /// Supported Visio editions (All/32bit/64bit)
        /// </summary>
        [Flags]
        private enum VisioEdition
        {
            X86,
            X64,
            All,
        };

        /// <summary>
        /// 
        /// </summary>
        [Flags]
        private enum AddonAttrs
        {
            PerformsActions = 1,
            HasAboutBox = 2,
            ProvidesHelp = 4,
            DisplayWaitCursor = 8,
            HideInUI = 16,
        };

        /// <summary>
        /// 
        /// </summary>
        [Flags]
        private enum EnablingPolicy
        {
            None = 0,

            AlwaysEnabled = 0xffff,
            DynamicallyEnabled = 0x0,
            StaticallyEnabled = 0x1,
            StaticallyThenDynamicallyEnabled = 0x8001,

            Document = 0x1,
            DocumentWindow = 0x2,
            DrawingWindow = 0x7,
            PageWindow = 0x87,
            MasterWindow = 0x107,
            StencilWindow = 0x0b,
            SheetWindow = 0x13,
            IconWindow = 0x23,
            TargetContext = 0x41,
            TargetContextPage = 0xC1,
            TargetContextMaster = 0x141,
            TargetContextSelection = 0x241,
        };



        /// <summary>
        /// Parses a "Visio" element.
        /// </summary>
        /// <param name="node">Element to parse.</param>
        /// <param name="parentNode">Parent element (File)</param>
        /// <param name="contextComponentId">Identifier for parent component.</param>
        /// <param name="visioContentType"></param>
        private void ParseVisioElement(XmlNode parentNode, XmlNode node, string contextComponentId, VisioContentType visioContentType)
        {
            var sourceLineNumbers = Preprocessor.GetSourceLineNumbers(node);

            if (parentNode.Attributes == null || parentNode.Attributes["Name"] == null)
            {
                Core.OnMessage(VisioErrors.FileHasNoAttributes(sourceLineNumbers));
                return;
            }

            var attribFileName = parentNode.Attributes["Name"];
            
            var fileName = 
                Core.GetAttributeLongFilename(sourceLineNumbers, attribFileName, false);

            string component = null;
            string feature = null;

            if (!ValidateVisioContentType(sourceLineNumbers, fileName, visioContentType))
            {
                return;
            }

            var visioVersion = GetDefaultVisioVersion(fileName);
            var visioEdition = VisioEdition.All;

            var addonAttrs = AddonAttrs.PerformsActions;

            string menuPath = Path.GetFileName(fileName);
            string altNames = "";

            var visioLanguage = 1;

            uint quickShapesCount = 0;
            bool featuredTemplate = true;

            string localizedName = Path.GetFileName(fileName);
            string universalName = Path.GetFileName(fileName);

            uint ordinal = 1;

            var enablingPolicy = EnablingPolicy.AlwaysEnabled;
            var staticEnableConditions = EnablingPolicy.None;
            bool invokeOnStartup = false;

            if (node.Attributes != null)
            foreach (XmlAttribute attrib in node.Attributes)
            {
                switch (attrib.LocalName)
                {
                    case "Feature":
                        feature = Core.GetAttributeIdentifierValue(sourceLineNumbers, attrib);
                        Core.CreateWixSimpleReferenceRow(sourceLineNumbers, "Feature", feature);
                        break;

                    case "Component":
                        component = Core.GetAttributeIdentifierValue(sourceLineNumbers, attrib);
                        Core.CreateWixSimpleReferenceRow(sourceLineNumbers, "Component", component);
                        break;

                    case "Visio2003":
                        if (YesNoType.Yes == Core.GetAttributeYesNoValue(sourceLineNumbers, attrib))
                            visioVersion |= VisioVersion.Visio2003;
                        else
                            visioVersion &= ~VisioVersion.Visio2003;
                        break;

                    case "Visio2007":
                        if (YesNoType.Yes == Core.GetAttributeYesNoValue(sourceLineNumbers, attrib))
                            visioVersion |= VisioVersion.Visio2007;
                        else
                            visioVersion &= ~VisioVersion.Visio2007;
                        break;

                    case "Visio2010":
                        if (YesNoType.Yes == Core.GetAttributeYesNoValue(sourceLineNumbers, attrib))
                            visioVersion |= VisioVersion.Visio2010;
                        else
                            visioVersion &= ~VisioVersion.Visio2010;
                        break;

                    case "Visio2013":
                        if (YesNoType.Yes == Core.GetAttributeYesNoValue(sourceLineNumbers, attrib))
                            visioVersion |= VisioVersion.Visio2013;
                        else
                            visioVersion &= ~VisioVersion.Visio2013;
                        break;

                    case "VisioEdition":
                        visioEdition = ParseVisioEditionAttributeValue(sourceLineNumbers, attrib);
                        break;

                    case "Language":
                        visioLanguage = ParseVisioLanguageCode(sourceLineNumbers, attrib);
                        break;

                    case "MenuPath":
                        if (visioContentType == VisioContentType.Template || visioContentType == VisioContentType.Stencil || visioContentType == VisioContentType.Addon)
                            menuPath = Core.GetAttributeBundleVariableValue(sourceLineNumbers, attrib);
                        else
                            Core.UnexpectedAttribute(sourceLineNumbers, attrib);
                        break;

                    case "AltNames":
                        if (visioContentType == VisioContentType.Template || visioContentType == VisioContentType.Stencil)
                            altNames = Core.GetAttributeBundleVariableValue(sourceLineNumbers, attrib);
                        else
                            Core.UnexpectedAttribute(sourceLineNumbers, attrib);
                        break;



                    case "QuickShapeCount":
                        if (visioContentType == VisioContentType.Stencil)
                            quickShapesCount = ParseUint(sourceLineNumbers, attrib);
                        else
                            Core.UnexpectedAttribute(sourceLineNumbers, attrib);
                        break;



                    case "FeaturedTemplate":
                        if (visioContentType == VisioContentType.Template)
                            featuredTemplate = (YesNoType.Yes == Core.GetAttributeYesNoValue(sourceLineNumbers, attrib));
                        else
                            Core.UnexpectedAttribute(sourceLineNumbers, attrib);
                        break;



                    case "LocalizedName":
                        if (visioContentType == VisioContentType.Addon)
                            localizedName = Core.GetAttributeBundleVariableValue(sourceLineNumbers, attrib);
                        else
                            Core.UnexpectedAttribute(sourceLineNumbers, attrib);
                        break;

                    case "UniversalName":
                        if (visioContentType == VisioContentType.Addon)
                            universalName = Core.GetAttributeBundleVariableValue(sourceLineNumbers, attrib);
                        else
                            Core.UnexpectedAttribute(sourceLineNumbers, attrib);
                        break;

                    case "Ordinal":
                        if (visioContentType == VisioContentType.Addon)
                            ordinal = ParseUint(sourceLineNumbers, attrib);
                        else
                            Core.UnexpectedAttribute(sourceLineNumbers, attrib);
                        break;


                    case "PerformsActions":
                        if (visioContentType != VisioContentType.Addon)
                            Core.UnexpectedAttribute(sourceLineNumbers, attrib);
                        else
                        {
                            if (YesNoType.Yes == Core.GetAttributeYesNoValue(sourceLineNumbers, attrib))
                                addonAttrs |= AddonAttrs.PerformsActions;
                            else
                                addonAttrs &= ~AddonAttrs.PerformsActions;
                        }
                        break;

                    case "HasAboutBox":
                        if (visioContentType != VisioContentType.Addon)
                            Core.UnexpectedAttribute(sourceLineNumbers, attrib);
                        else
                        {
                            if (YesNoType.Yes == Core.GetAttributeYesNoValue(sourceLineNumbers, attrib))
                                addonAttrs |= AddonAttrs.HasAboutBox;
                            else
                                addonAttrs &= ~AddonAttrs.HasAboutBox;
                        }
                        break;

                    case "ProvidesHelp":
                        if (visioContentType != VisioContentType.Addon)
                            Core.UnexpectedAttribute(sourceLineNumbers, attrib);
                        else
                        {
                            if (YesNoType.Yes == Core.GetAttributeYesNoValue(sourceLineNumbers, attrib))
                                addonAttrs |= AddonAttrs.ProvidesHelp;
                            else
                                addonAttrs &= ~AddonAttrs.ProvidesHelp;
                        }
                        break;

                    case "DisplayWaitCursor":
                        if (visioContentType != VisioContentType.Addon)
                            Core.UnexpectedAttribute(sourceLineNumbers, attrib);
                        else
                        {
                            if (YesNoType.Yes == Core.GetAttributeYesNoValue(sourceLineNumbers, attrib))
                                addonAttrs |= AddonAttrs.DisplayWaitCursor;
                            else
                                addonAttrs &= ~AddonAttrs.DisplayWaitCursor;
                        }
                        break;

                    case "HideInUI":
                        if (visioContentType != VisioContentType.Addon)
                            Core.UnexpectedAttribute(sourceLineNumbers, attrib);
                        else
                        {
                            if (YesNoType.Yes == Core.GetAttributeYesNoValue(sourceLineNumbers, attrib))
                                addonAttrs |= AddonAttrs.HideInUI;
                            else
                                addonAttrs &= ~AddonAttrs.HideInUI;
                        }
                        break;

                    case "EnablingPolicy":
                        if (visioContentType != VisioContentType.Addon)
                            Core.UnexpectedAttribute(sourceLineNumbers, attrib);
                        else
                        {
                            enablingPolicy = ParseEnablingPolicy(sourceLineNumbers, attrib);
                        }
                        break;

                    case "StaticEnableConditions":
                        if (visioContentType != VisioContentType.Addon)
                            Core.UnexpectedAttribute(sourceLineNumbers, attrib);
                        else
                        {
                            staticEnableConditions = ParseStaticEnableConditions(sourceLineNumbers, attrib);
                        }
                        break;

                    case "InvokeOnStartup":
                        if (visioContentType != VisioContentType.Addon)
                            Core.UnexpectedAttribute(sourceLineNumbers, attrib);
                        else
                        {
                            invokeOnStartup = (YesNoType.Yes == Core.GetAttributeYesNoValue(sourceLineNumbers, attrib));
                        }
                        break;

                    default:
                        Core.UnexpectedAttribute(sourceLineNumbers, attrib);
                        break;
                }
            }

            if (string.IsNullOrEmpty(menuPath))
            {
                if (visioContentType == VisioContentType.Template || visioContentType == VisioContentType.Stencil)
                    Core.OnMessage(WixErrors.ExpectedAttribute(sourceLineNumbers, node.Name, "MenuPath"));
            }

            if ((visioEdition & VisioEdition.X86) == VisioEdition.X86)
                Core.CreateWixSimpleReferenceRow(sourceLineNumbers, "CustomAction", "SetConfigChangeID");

            if ((visioEdition & VisioEdition.X64) == VisioEdition.X64)
                Core.CreateWixSimpleReferenceRow(sourceLineNumbers, "CustomAction", "SetConfigChangeID64");

            if (Core.EncounteredError)
                return;

            var featureId = feature ?? Guid.Empty.ToString("B");
            string componentId = component ?? contextComponentId;
            string editionCode = GetVisioEditionCode(visioEdition);

            var enablingPolicyCode = (int) (enablingPolicy | staticEnableConditions);
            var addonAttrsCode = (int) (addonAttrs);
            int invokeOnStartupCode = (invokeOnStartup ? 1 : 0);

            if ((visioVersion & VisioVersion.Visio2003) == VisioVersion.Visio2003)
            {
                string appData2003 = "";
                    
                switch (visioContentType)
                {
                    case VisioContentType.Stencil:
                    case VisioContentType.Template:
                        appData2003 = string.Format(@"{0}|{1}", menuPath, altNames);
                        break;

                    case VisioContentType.Help:
                        appData2003 = null;
                        break;

                    case VisioContentType.Addon:
                        appData2003 = string.Format(@"{0}|{1}|{2}|{3}|{4}|{5}|{6}",
                            menuPath, localizedName, universalName, ordinal, addonAttrsCode, enablingPolicyCode, invokeOnStartupCode);
                        break;
                }

                if (visioLanguage == 1)
                {
                    foreach (var lcid in GetAllVisioLanguageCodes())
                    {
                        var qualifier2003 = MakeQualifier(visioContentType, lcid, fileName, ordinal);

                        GenerateRow(sourceLineNumbers, GetVisioComponentId(visioContentType, VisioVersion.Visio2003),
                                    qualifier2003, appData2003, featureId, componentId);
                    }
                }
                else
                {
                    string qualifier2003 = MakeQualifier(visioContentType, visioLanguage, fileName, ordinal); 
                        
                    GenerateRow(sourceLineNumbers, GetVisioComponentId(visioContentType, VisioVersion.Visio2003),
                                qualifier2003, appData2003, featureId, componentId);
                }
            }

            string qualifier = MakeQualifier(visioContentType, visioLanguage, fileName, ordinal);

            if ((visioVersion & VisioVersion.Visio2007) == VisioVersion.Visio2007)
            {
                string appData2007 = "";
                    
                switch (visioContentType)
                {
                    case VisioContentType.Stencil:
                        appData2007 = string.Format(@"{0}|{1}", menuPath, altNames);
                        break;

                    case VisioContentType.Template:
                        appData2007 = string.Format(@"{0}|{1}|{2}", menuPath, altNames, featuredTemplate ? 1 : 0);
                        break;

                    case VisioContentType.Help:
                        appData2007 = "-1";
                        break;

                    case VisioContentType.Addon:
                        appData2007 = string.Format(@"{0}|{1}|{2}|{3}|{4}|{5}|{6}",
                            menuPath, localizedName, universalName, ordinal, addonAttrsCode, enablingPolicyCode, invokeOnStartupCode);
                        break;
                }

                GenerateRow(sourceLineNumbers, GetVisioComponentId(visioContentType, VisioVersion.Visio2007),
                            qualifier, appData2007, featureId, componentId);
                    
            }
            if ((visioVersion & VisioVersion.Visio2010) == VisioVersion.Visio2010)
            {
                string appData2010 = "";
                switch (visioContentType)
                {
                    case VisioContentType.Stencil:  
                        appData2010 = string.Format(@"{0}|{1}|{2}|{3}", menuPath, altNames, quickShapesCount, editionCode);
                        break;

                    case VisioContentType.Template: 
                        appData2010 = string.Format(@"{0}|{1}|{2}|{3}", menuPath, altNames, 1, editionCode);
                        break;

                    case VisioContentType.Help:
                        appData2010 = "-1";
                        break;

                    case VisioContentType.Addon:
                        appData2010 = string.Format(@"{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}",
                            menuPath, localizedName, universalName, ordinal, addonAttrsCode, enablingPolicyCode, invokeOnStartupCode, editionCode);
                        break;

                }
                    
                GenerateRow(sourceLineNumbers, GetVisioComponentId(visioContentType, VisioVersion.Visio2010),
                            qualifier, appData2010, featureId, componentId);
            }

            if ((visioVersion & VisioVersion.Visio2013) == VisioVersion.Visio2013)
            {
                string appData2013 = "";

                switch (visioContentType)
                {
                    case VisioContentType.Stencil:
                        appData2013 = string.Format(@"{0}|{1}|{2}|{3}", menuPath, altNames, quickShapesCount, editionCode);
                        break;

                    case VisioContentType.Template:
                        appData2013 = string.Format(@"{0}|{1}|{2}|{3}", menuPath, altNames, 1, editionCode);
                        break;

                    case VisioContentType.Help:
                        appData2013 = "-1";
                        break;

                    case VisioContentType.Addon:
                        appData2013 = string.Format(@"{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}",
                            menuPath, localizedName, universalName, ordinal, addonAttrsCode, enablingPolicyCode, invokeOnStartupCode, editionCode);
                        break;
                }

                GenerateRow(sourceLineNumbers, GetVisioComponentId(visioContentType, VisioVersion.Visio2013),
                            qualifier, appData2013, featureId, componentId);
            }
        }

        private static string MakeQualifier(VisioContentType visioContentType, int lcid, string fileName, uint ordinal)
        {
            return (visioContentType == VisioContentType.Addon)
                ? string.Format(@"{0}\{1}\{2}", lcid, ordinal - 1, fileName)
                : string.Format(@"{0}\{1}", lcid, fileName);
        }


        private void GenerateRow(SourceLineNumberCollection sourceLineNumbers, string visioComponentId, string qualifier,
                                 string appData, string featureId, string componentId)
        {
            Row row = Core.CreateRow(sourceLineNumbers, "PublishComponent");

            row[0] = visioComponentId;
            row[1] = qualifier;
            row[2] = componentId;
            row[3] = appData;
            row[4] = featureId;
        }



        private VisioEdition ParseVisioEditionAttributeValue(SourceLineNumberCollection sourceLineNumbers, XmlAttribute attrib)
        {
            string attribValue = Core.GetAttributeBundleVariableValue(sourceLineNumbers, attrib);
            switch (attribValue)
            {
                case "x86":
                    return VisioEdition.X86;

                case "x64":
                    return VisioEdition.X64;

                default:
                    Core.OnMessage(VisioErrors.InvalidVisioEdition(sourceLineNumbers, attribValue));
                    return VisioEdition.All;
            }
        }


        private EnablingPolicy ParseEnablingPolicy(SourceLineNumberCollection sourceLineNumbers, XmlAttribute attrib)
        {
            string attribValue = Core.GetAttributeBundleVariableValue(sourceLineNumbers, attrib);
            switch (attribValue)
            {
                case "AlwaysEnabled":
                    return EnablingPolicy.AlwaysEnabled;

                case "DynamicallyEnabled":
                    return EnablingPolicy.DynamicallyEnabled;

                case "StaticallyEnabled":
                    return EnablingPolicy.StaticallyEnabled;

                case "StaticallyThenDynamicallyEnabled":
                    return EnablingPolicy.StaticallyThenDynamicallyEnabled;

                default:
                    Core.OnMessage(VisioErrors.InvalidVisioEdition(sourceLineNumbers, attribValue));
                    return EnablingPolicy.AlwaysEnabled;
            }
        }

        private EnablingPolicy ParseStaticEnableConditions(SourceLineNumberCollection sourceLineNumbers, XmlAttribute attrib)
        {
            string attribValue = Core.GetAttributeBundleVariableValue(sourceLineNumbers, attrib);
            switch (attribValue)
            {
                case "Document":
                    return EnablingPolicy.Document;

                case "DocumentWindow":
                    return EnablingPolicy.DocumentWindow;

                case "DrawingWindow":
                    return EnablingPolicy.DrawingWindow;

                case "PageWindow":
                    return EnablingPolicy.PageWindow;

                case "MasterWindow":
                    return EnablingPolicy.MasterWindow;

                case "StencilWindow":
                    return EnablingPolicy.StencilWindow;

                case "SheetWindow":
                    return EnablingPolicy.SheetWindow;

                case "IconWindow":
                    return EnablingPolicy.IconWindow;

                case "TargetContext":
                    return EnablingPolicy.TargetContext;

                case "TargetContextPage":
                    return EnablingPolicy.TargetContextPage;

                case "TargetContextMaster":
                    return EnablingPolicy.TargetContextMaster;

                case "TargetContextSelection":
                    return EnablingPolicy.TargetContextSelection;
                    

                default:
                    Core.OnMessage(VisioErrors.InvalidVisioEdition(sourceLineNumbers, attribValue));
                    return EnablingPolicy.Document;
            }
        }
        
        private static string GetVisioEditionCode(VisioEdition visioEdition)
        {
            switch (visioEdition)
            {
                case VisioEdition.All:
                    return "-1";
                case VisioEdition.X86:
                    return "32";
                case VisioEdition.X64:
                    return "64";

                default:
                    throw new ArgumentOutOfRangeException("visioEdition");
            }
        }

        /// <summary>
        /// Type of Visio content to register: template/stencil/etc
        /// </summary>
        private enum VisioContentType
        {
            Template = 0,
            Stencil = 1,
            Help = 2,
            Addon = 3,
        };

        /// <summary>
        /// Returns default set of Visio versions to install. By Default, it is 2007/2010/2013.
        /// For new file types (2013) the default is to publish for Visio 2013 only.
        /// </summary>
        /// <param name="fileName">File name to analyze</param>
        /// <returns>Set of Visio version to publish for by default</returns>
        private VisioVersion GetDefaultVisioVersion(string fileName)
        {
            switch (Path.GetExtension(fileName))
            {
                case ".chm":
                case ".vsl":
                case ".exe":
                case ".vss":
                case ".vst":
                case ".vsx":
                case ".vtx":
                    return VisioVersion.Visio2007 | VisioVersion.Visio2010 | VisioVersion.Visio2013;

                case ".vssx":
                case ".vstx":
                case ".vstm":
                    return VisioVersion.Visio2013;

                default:
                    return VisioVersion.Visio2003;
            }
        }

        /// <summary>
        /// Detects Visio content type based on file extension.
        /// Reports error if file under which the publish component was located is not actually Visio file.
        /// </summary>
        /// <param name="sourceLineNumbers"></param>
        /// <param name="fileName"></param>
        /// <param name="visioContentType"></param>
        /// <returns></returns>
        private bool ValidateVisioContentType(SourceLineNumberCollection sourceLineNumbers, string fileName, VisioContentType visioContentType)
        {
            var extension = Path.GetExtension(fileName);
            switch (extension)
            {
                case ".vss":
                case ".vsx":
                case ".vssx":
                case ".vssm":
                    if (visioContentType != VisioContentType.Stencil)
                    {
                        Core.OnMessage(VisioErrors.InvalidStencilFileExtension(sourceLineNumbers, extension));
                        return false;
                    }
                    break;

                case ".vst":
                case ".vtx":
                case ".vstx":
                case ".vstm":
                    if (visioContentType != VisioContentType.Template)
                    {
                        Core.OnMessage(VisioErrors.InvalidTemplateFileExtension(sourceLineNumbers, extension));
                        return false;
                    }
                    break;

                case ".vsl":
                case ".exe":
                    if (visioContentType != VisioContentType.Addon)
                    {
                        Core.OnMessage(VisioErrors.InvalidAddonFileExtension(sourceLineNumbers, extension));
                        return false;
                    }
                    break;

                case ".chm":
                    if (visioContentType != VisioContentType.Help)
                    {
                        Core.OnMessage(VisioErrors.InvalidHelpFileExtension(sourceLineNumbers, extension));
                        return false;
                    }
                    break;

                default:
                    Core.OnMessage(VisioErrors.InvalidFileExtension(sourceLineNumbers, extension));
                    return false;
            }

            return true;
        }

        /// <summary>
        /// Returns the ComponentID for PublishComponent table for given combination of version/content type.
        /// </summary>
        /// <param name="contentType"></param>
        /// <param name="visioVersion"></param>
        /// <returns></returns>

        private string GetVisioComponentId(VisioContentType contentType, VisioVersion visioVersion)
        {
            var pattern = GetVisioBaseComponentId(visioVersion);

            return pattern.Replace("X", Convert.ToInt32(contentType).ToString(CultureInfo.InvariantCulture));
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="visioVersion"></param>
        /// <returns></returns>

        private string GetVisioBaseComponentId(VisioVersion visioVersion)
        {
            switch (visioVersion)
            {
                case VisioVersion.Visio2003:
                    return "{CF1F488D-8D6F-499C-A78D-026E1DF3810X}";

                case VisioVersion.Visio2007:
                    return "{6D9D8B6F-D0EF-4BC0-8DD4-09DD6CE2B00X}";

                case VisioVersion.Visio2010:
                    return "{6D9D8B6F-D0EF-4BC0-8DD4-09DD6CE2B10X}";

                case VisioVersion.Visio2013:
                    return "{6D9D8B6F-D0EF-4BC0-8DD4-09DD6CE2B20X}";

                default:
                    throw new ArgumentOutOfRangeException("visioVersion");
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sourceLineNumbers"></param>
        /// <param name="attrib"></param>
        /// <returns></returns>

        private uint ParseUint(SourceLineNumberCollection sourceLineNumbers, XmlAttribute attrib)
        {
            string attribValue = Core.GetAttributeBundleVariableValue(sourceLineNumbers, attrib);

            uint result;
            if (uint.TryParse(attribValue, out result))
                return result;

            Core.OnMessage(VisioErrors.InvalidUint(sourceLineNumbers, attrib.LocalName, attribValue));

            return 0;
        }

        /// <summary>
        /// Parses Language attribute into LCID code.
        /// Language attribute can be either a LCID directly (in this case it is verified that it exists) or language name,
        /// in this case it is converted to LCID.
        /// </summary>
        /// <param name="sourceLineNumbers">Source line number for the parent element.</param>
        /// <param name="attrib">Language attribute to parse</param>
        /// <returns>Language code (LCID)</returns>
        /// 
        private int ParseVisioLanguageCode(SourceLineNumberCollection sourceLineNumbers, XmlAttribute attrib)
        {
            var attribValue = Core.GetAttributeBundleVariableValue(sourceLineNumbers, attrib);

            try
            {
                int lcid;
                var cultureInfo = int.TryParse(attribValue, out lcid) 
                    ? CultureInfo.GetCultureInfo(lcid) 
                    : CultureInfo.GetCultureInfoByIetfLanguageTag(attribValue);

                return cultureInfo.LCID;
            }
            catch (Exception)
            {
                Core.OnMessage(VisioErrors.InvalidLanguage(sourceLineNumbers, attribValue));
                return 0;
            }
        }

        /// <summary>
        /// List of all languages supported by Visio 2003. We need to add those explicitly for Visio 2003
        /// </summary>
        /// <returns></returns>
        private IEnumerable<int> GetAllVisioLanguageCodes()
        {
            return new[]
                       {
                           1025, 1028, 1029, 1030, 
                           1031, 1032, 1033, 1035, 
                           1036, 1037, 1038, 1040, 
                           1041, 1042, 1043, 1044, 
                           1045, 1046, 1048, 1049, 
                           1051, 1053, 1055, 1058, 
                           1060, 2052, 2070, 3082
                       };

        }

    }
}