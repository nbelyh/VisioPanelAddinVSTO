using System;
using System.IO;
using System.Reflection;
using System.Xml;
using System.Xml.Schema;
using Microsoft.Tools.WindowsInstallerXml;
using System.Globalization;

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
                    
                    string fileId = contextValues[0];
                    string componentId = contextValues[1];

                    Func<string> fileName = () => 
                        parentElement.HasAttribute("Name")
                        ? Core.GetAttributeLongFilename(sourceLineNumbers, parentElement.Attributes["Name"], false)
                        : parentElement.HasAttribute("Source")
                            ? Path.GetFileNameWithoutExtension(Core.GetAttributeValue(sourceLineNumbers, parentElement.Attributes["Source"], false))
                            : null;

                    switch (element.LocalName)
                    {
                        case "PublishStencil":
                            ParseVisioElement(element, componentId, fileName(), VisioContentType.Stencil);
                            break;

                        case "PublishTemplate":
                            ParseVisioElement(element, componentId, fileName(), VisioContentType.Template);
                            break;

                        case "PublishHelpFile":
                            ParseVisioElement(element, componentId, fileName(), VisioContentType.Help);
                            break;

                        case "PublishAddon":
                            ParseVisioElement(element, componentId, fileName(), VisioContentType.Addon);
                            break;

                        case "Publish":
                            Core.OnMessage(VisioErrors.UsingPublish(sourceLineNumbers));
                            break;

                        case "PublishAddinVSTO":
                            ParseAddinElement(element, fileName(), fileId);
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

        private void ParseAddinElement(XmlElement node, string friendlyName, string fileId)
        {
            string description = null;

            var visioEdition = VisioEdition.Default;
            string id = null;
            var sourceLineNumbers = Preprocessor.GetSourceLineNumbers(node);
            var commandLinkeSafe = 1;
            var loadBehavior = 3;

            foreach (XmlAttribute attrib in node.Attributes)
            {
                switch (attrib.LocalName)
                {
                    case "File":
                        fileId = Core.GetAttributeIdentifierValue(sourceLineNumbers, attrib);
                        Core.CreateWixSimpleReferenceRow(sourceLineNumbers, "File", fileId);
                        break;

                    case "Id":
                        id = Core.GetAttributeIdentifierValue(sourceLineNumbers, attrib);
                        break;

                    case "FriendlyName":
                        friendlyName = Core.GetAttributeBundleVariableValue(sourceLineNumbers, attrib);
                        break;

                    case "Description":
                        description = Core.GetAttributeBundleVariableValue(sourceLineNumbers, attrib);
                        break;

                    case "VisioEdition":
                        visioEdition = ParseVisioEditionAttributeValue(sourceLineNumbers, attrib);
                        break;

                    case "CommandLineSafe":
                        commandLinkeSafe = (YesNoType.Yes == Core.GetAttributeYesNoValue(sourceLineNumbers, attrib) ? 1 : 0);
                        break;

                    case "LoadBehavior":
                        loadBehavior = Core.GetAttributeIntegerValue(sourceLineNumbers, attrib, 0, 32);
                        break;

                    default:
                        Core.UnexpectedAttribute(sourceLineNumbers, attrib);
                        break;
                }
            }

            if (Core.EncounteredError)
                return;

            Core.EnsureTable(sourceLineNumbers, "AddinRegistration");
            var r = Core.CreateRow(sourceLineNumbers, "AddinRegistration");

            if (id == null)
                id = CompilerCore.GetIdentifierFromName(friendlyName);

            r[0] = id;
            r[1] = fileId;
            r[2] = friendlyName;
            r[3] = description;
            r[4] = (int)visioEdition;
            r[5] = commandLinkeSafe;
            r[6] = loadBehavior;

            Core.CreateWixSimpleReferenceRow(sourceLineNumbers, "CustomAction", "SchedAddinRegistration");
        }

        /// <summary>
        /// Parses a "Visio" element.
        /// </summary>
        /// <param name="node">Element to parse.</param>
        /// <param name="contextComponentId">Identifier for parent component.</param>
        /// <param name="fileName"></param>
        /// <param name="visioContentType"></param>
        private void ParseVisioElement(XmlElement node, string contextComponentId, string fileName, VisioContentType visioContentType)
        {
            var sourceLineNumbers = Preprocessor.GetSourceLineNumbers(node);

            string component = null;
            string feature = null;

            if (visioContentType != GetVisioContentType(fileName))
                Core.OnMessage(VisioWarnings.InvalidFileExtension(sourceLineNumbers, fileName, Enum.GetName(typeof(VisioContentType), visioContentType)));

            var publishInfo = new VisioPublishInfo(fileName);

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
                        ParseVisioVersionAttribute(sourceLineNumbers, attrib, publishInfo, VisioVersion.Visio2003);
                        break;

                    case "Visio2007":
                        ParseVisioVersionAttribute(sourceLineNumbers, attrib, publishInfo, VisioVersion.Visio2007);
                        break;

                    case "Visio2010":
                        ParseVisioVersionAttribute(sourceLineNumbers, attrib, publishInfo, VisioVersion.Visio2010);
                        break;

                    case "Visio2013":
                        ParseVisioVersionAttribute(sourceLineNumbers, attrib, publishInfo, VisioVersion.Visio2013);
                        break;

                    case "Visio2016":
                        ParseVisioVersionAttribute(sourceLineNumbers, attrib, publishInfo, VisioVersion.Visio2016);
                        break;

                    case "VisioEdition":
                        publishInfo.VisioEdition = ParseVisioEditionAttributeValue(sourceLineNumbers, attrib);
                        break;

                    case "Language":
                        publishInfo.VisioLanguage = ParseVisioLanguageCode(sourceLineNumbers, attrib);
                        break;

                    case "MenuPath":
                        if (visioContentType == VisioContentType.Template || visioContentType == VisioContentType.Stencil || visioContentType == VisioContentType.Addon)
                            publishInfo.MenuPath = Core.GetAttributeBundleVariableValue(sourceLineNumbers, attrib);
                        else
                            Core.UnexpectedAttribute(sourceLineNumbers, attrib);
                        break;

                    case "AltNames":
                        if (visioContentType == VisioContentType.Template || visioContentType == VisioContentType.Stencil)
                            publishInfo.AltNames = Core.GetAttributeBundleVariableValue(sourceLineNumbers, attrib);
                        else
                            Core.UnexpectedAttribute(sourceLineNumbers, attrib);
                        break;



                    case "QuickShapeCount":
                        if (visioContentType == VisioContentType.Stencil)
                            publishInfo.QuickShapesCount = ParseUint(sourceLineNumbers, attrib);
                        else
                            Core.UnexpectedAttribute(sourceLineNumbers, attrib);
                        break;



                    case "FeaturedTemplate":
                        if (visioContentType == VisioContentType.Template)
                            publishInfo.FeaturedTemplate = (YesNoType.Yes == Core.GetAttributeYesNoValue(sourceLineNumbers, attrib));
                        else
                            Core.UnexpectedAttribute(sourceLineNumbers, attrib);
                        break;



                    case "LocalizedName":
                        if (visioContentType == VisioContentType.Addon)
                            publishInfo.LocalizedName = Core.GetAttributeBundleVariableValue(sourceLineNumbers, attrib);
                        else
                            Core.UnexpectedAttribute(sourceLineNumbers, attrib);
                        break;

                    case "UniversalName":
                        if (visioContentType == VisioContentType.Addon)
                            publishInfo.UniversalName = Core.GetAttributeBundleVariableValue(sourceLineNumbers, attrib);
                        else
                            Core.UnexpectedAttribute(sourceLineNumbers, attrib);
                        break;

                    case "Ordinal":
                        if (visioContentType == VisioContentType.Addon)
                            publishInfo.Ordinal = ParseUint(sourceLineNumbers, attrib);
                        else
                            Core.UnexpectedAttribute(sourceLineNumbers, attrib);
                        break;


                    case "PerformsActions":
                        ParseAddonAttrs(visioContentType, sourceLineNumbers, attrib, publishInfo, AddonAttrs.PerformsActions);
                        break;

                    case "HasAboutBox":
                        ParseAddonAttrs(visioContentType, sourceLineNumbers, attrib, publishInfo, AddonAttrs.HasAboutBox);
                        break;

                    case "ProvidesHelp":
                        ParseAddonAttrs(visioContentType, sourceLineNumbers, attrib, publishInfo, AddonAttrs.ProvidesHelp);
                        break;

                    case "DisplayWaitCursor":
                        ParseAddonAttrs(visioContentType, sourceLineNumbers, attrib, publishInfo, AddonAttrs.DisplayWaitCursor);
                        break;

                    case "HideInUI":
                        ParseAddonAttrs(visioContentType, sourceLineNumbers, attrib, publishInfo, AddonAttrs.HideInUI);
                        break;

                    case "EnablingPolicy":
                        if (visioContentType == VisioContentType.Addon)
                            publishInfo.EnablingPolicy = ParseEnablingPolicy(sourceLineNumbers, attrib);
                        else
                            Core.UnexpectedAttribute(sourceLineNumbers, attrib);
                        break;

                    case "StaticEnableConditions":
                        if (visioContentType == VisioContentType.Addon)
                            publishInfo.StaticEnableConditions = ParseStaticEnableConditions(sourceLineNumbers, attrib);
                        else
                            Core.UnexpectedAttribute(sourceLineNumbers, attrib);
                        break;

                    case "InvokeOnStartup":
                        if (visioContentType == VisioContentType.Addon)
                            publishInfo.InvokeOnStartup = (YesNoType.Yes == Core.GetAttributeYesNoValue(sourceLineNumbers, attrib));
                        else
                            Core.UnexpectedAttribute(sourceLineNumbers, attrib);
                        break;

                    default:
                        Core.UnexpectedAttribute(sourceLineNumbers, attrib);
                        break;
                }
            }

            if (string.IsNullOrEmpty(publishInfo.MenuPath))
            {
                if (visioContentType == VisioContentType.Template || visioContentType == VisioContentType.Stencil)
                    Core.OnMessage(WixErrors.ExpectedAttribute(sourceLineNumbers, node.Name, "MenuPath"));
            }

            if (publishInfo.VisioVersion == VisioVersion.Default)
                publishInfo.VisioVersion = VisioPublishInfo.GetDefaultVisioVersion(fileName);

            if (publishInfo.VisioEdition == VisioEdition.Default)
                publishInfo.VisioEdition = (VisioEdition.X86 | VisioEdition.X64);

            if ((publishInfo.VisioEdition & VisioEdition.X86) == VisioEdition.X86)
                Core.CreateWixSimpleReferenceRow(sourceLineNumbers, "CustomAction", "UpdateConfigChangeID32");

            if ((publishInfo.VisioEdition & VisioEdition.X64) == VisioEdition.X64)
                Core.CreateWixSimpleReferenceRow(sourceLineNumbers, "CustomAction", "UpdateConfigChangeID64");

            if (Core.EncounteredError)
                return;

            var featureId = feature ?? Guid.Empty.ToString("B");
            var componentId = component ?? contextComponentId;

            foreach (var rowInfo in publishInfo.GenerateRowInfos(visioContentType, fileName))
            {
                var row = Core.CreateRow(sourceLineNumbers, "PublishComponent");

                row[0] = rowInfo.VisioComponentId;
                row[1] = rowInfo.Qualifier;
                row[2] = componentId;
                row[3] = rowInfo.AppData;
                row[4] = featureId;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="visioContentType"></param>
        /// <param name="sourceLineNumbers"></param>
        /// <param name="attrib"></param>
        /// <param name="publishInfo"></param>
        /// <param name="addonAttrs"></param>

        private void ParseAddonAttrs(VisioContentType visioContentType, SourceLineNumberCollection sourceLineNumbers,
                                     XmlAttribute attrib, VisioPublishInfo publishInfo, AddonAttrs addonAttrs)
        {
            if (visioContentType != VisioContentType.Addon)
                Core.UnexpectedAttribute(sourceLineNumbers, attrib);
            else
            {
                if (YesNoType.Yes == Core.GetAttributeYesNoValue(sourceLineNumbers, attrib))
                    publishInfo.AddonAttrs |= addonAttrs;
                else
                    publishInfo.AddonAttrs &= ~addonAttrs;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sourceLineNumbers"></param>
        /// <param name="attrib"></param>
        /// <param name="publishInfo"></param>
        /// <param name="visioVersion"></param>

        private void ParseVisioVersionAttribute(SourceLineNumberCollection sourceLineNumbers, XmlAttribute attrib,
                                                VisioPublishInfo publishInfo, VisioVersion visioVersion)
        {
            if (YesNoType.Yes == Core.GetAttributeYesNoValue(sourceLineNumbers, attrib))
                publishInfo.VisioVersion |= visioVersion;
            else
                publishInfo.VisioVersion &= ~visioVersion;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sourceLineNumbers"></param>
        /// <param name="attrib"></param>
        /// <returns></returns>

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
                    return VisioEdition.Default;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sourceLineNumbers"></param>
        /// <param name="attrib"></param>
        /// <returns></returns>

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

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sourceLineNumbers"></param>
        /// <param name="attrib"></param>
        /// <returns></returns>

        private StaticEnableConditions ParseStaticEnableConditions(SourceLineNumberCollection sourceLineNumbers, XmlAttribute attrib)
        {
            string attribValue = Core.GetAttributeBundleVariableValue(sourceLineNumbers, attrib);
            switch (attribValue)
            {
                case "Document":
                    return StaticEnableConditions.Document;

                case "DocumentWindow":
                    return StaticEnableConditions.DocumentWindow;

                case "DrawingWindow":
                    return StaticEnableConditions.DrawingWindow;

                case "PageWindow":
                    return StaticEnableConditions.PageWindow;

                case "MasterWindow":
                    return StaticEnableConditions.MasterWindow;

                case "StencilWindow":
                    return StaticEnableConditions.StencilWindow;

                case "SheetWindow":
                    return StaticEnableConditions.SheetWindow;

                case "IconWindow":
                    return StaticEnableConditions.IconWindow;

                case "TargetContext":
                    return StaticEnableConditions.TargetContext;

                case "TargetContextPage":
                    return StaticEnableConditions.TargetContextPage;

                case "TargetContextMaster":
                    return StaticEnableConditions.TargetContextMaster;

                case "TargetContextSelection":
                    return StaticEnableConditions.TargetContextSelection;
                    

                default:
                    Core.OnMessage(VisioErrors.InvalidVisioEdition(sourceLineNumbers, attribValue));
                    return StaticEnableConditions.Document;
            }
        }

        /// <summary>
        /// Detects Visio content type based on file extension.
        /// Reports error if file under which the publish component was located is not actually Visio file.
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        private static VisioContentType GetVisioContentType(string fileName)
        {
            switch (Path.GetExtension(fileName))
            {
                case ".vss":
                case ".vsx":
                case ".vssx":
                case ".vssm":
                    return VisioContentType.Stencil;

                case ".vst":
                case ".vtx":
                case ".vstx":
                case ".vstm":
                    return VisioContentType.Template;

                case ".vsl":
                case ".exe":
                    return VisioContentType.Addon;

                case ".chm":
                    return VisioContentType.Help;

                default:
                    return VisioContentType.Unknown;
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

    }
}