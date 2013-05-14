using System;
using System.Collections.Generic;
using System.Reflection;
using System.Xml;
using System.Xml.Schema;
using Microsoft.Tools.WindowsInstallerXml;
using System.IO;
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
                    string componentId = contextValues[1];

                    switch (element.LocalName)
                    {
                        case "Publish":
                            ParseVisioElement(parentElement, element, componentId);
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
        /// Parses a "Visio" element.
        /// </summary>
        /// <param name="node">Element to parse.</param>
        /// <param name="parentNode">Parent element (File)</param>
        /// <param name="contextComponentId">Identifier for parent component.</param>
        private void ParseVisioElement(XmlNode parentNode, XmlNode node, string contextComponentId)
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

            var visioContentType = GetVisioContentType(sourceLineNumbers, fileName);
            var visioVersion = GetDefaultVisioVersion(fileName);
            var visioEdition = VisioEdition.All;

            string menuPath = null;
            string altNames = "";

            var visioLanguage = 1;

            uint quickShapesCount = 0;
            if (visioContentType == VisioContentType.Template)
                quickShapesCount = 1;

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
                        menuPath = ParseMenuString(sourceLineNumbers, attrib);
                        break;

                    case "AltNames":
                        altNames = ParseAltNames(sourceLineNumbers, attrib);
                        break;

                    case "QuickShapeCount":
                        quickShapesCount = ParseQuickShapeCount(sourceLineNumbers, attrib);
                        break;

                    default:
                        Core.UnexpectedAttribute(sourceLineNumbers, attrib);
                        break;
                }
            }

            if (menuPath == null)
                Core.OnMessage(WixErrors.ExpectedAttribute(sourceLineNumbers, node.Name, "MenuPath"));

            if ((visioEdition & VisioEdition.X86) == VisioEdition.X86)
                Core.CreateWixSimpleReferenceRow(sourceLineNumbers, "CustomAction", "SetConfigChangeID");

            if ((visioEdition & VisioEdition.X64) == VisioEdition.X64)
                Core.CreateWixSimpleReferenceRow(sourceLineNumbers, "CustomAction", "SetConfigChangeID64");

            if (!Core.EncounteredError)
            {
                var featureId = feature ?? Guid.Empty.ToString("B");
                string componentId = component ?? contextComponentId;
                string editionCode = GetVisioEditionCode(visioEdition);

                if ((visioVersion & VisioVersion.Visio2003) == VisioVersion.Visio2003)
                {
                    var appData2003 = string.Format(@"{0}|{1}", menuPath, altNames);
                    if (visioLanguage == 1)
                    {
                        foreach (var lcid in GetAllVisioLanguageCodes())
                        {
                            string qualifier2003 = string.Format(@"{0}\{1}", lcid, fileName);

                            GenerateRow(sourceLineNumbers, GetVisioComponentId(visioContentType, VisioVersion.Visio2003),
                                        qualifier2003, appData2003, featureId, componentId);
                        }
                    }
                    else
                    {
                        string qualifier2003 = string.Format(@"{0}\{1}", visioLanguage, fileName);

                        GenerateRow(sourceLineNumbers, GetVisioComponentId(visioContentType, VisioVersion.Visio2003),
                                    qualifier2003, appData2003, featureId, componentId);
                    }
                }

                string qualifier = string.Format(@"{0}\{1}", visioLanguage, fileName);

                if ((visioVersion & VisioVersion.Visio2007) == VisioVersion.Visio2007)
                {
                    string appData2007 = string.Format(@"{0}|{1}", menuPath, altNames);
                    GenerateRow(sourceLineNumbers, GetVisioComponentId(visioContentType, VisioVersion.Visio2007),
                                qualifier, appData2007, featureId, componentId);
                    
                }
                if ((visioVersion & VisioVersion.Visio2010) == VisioVersion.Visio2010)
                {
                    string appData2010 = string.Format(@"{0}|{1}|{2}|{3}", menuPath, altNames, quickShapesCount, editionCode);
                    GenerateRow(sourceLineNumbers, GetVisioComponentId(visioContentType, VisioVersion.Visio2010),
                                qualifier, appData2010, featureId, componentId);
                }

                if ((visioVersion & VisioVersion.Visio2013) == VisioVersion.Visio2013)
                {
                    string appData2013 = string.Format(@"{0}|{1}|{2}|{3}", menuPath, altNames, quickShapesCount, editionCode);
                    GenerateRow(sourceLineNumbers, GetVisioComponentId(visioContentType, VisioVersion.Visio2013),
                                qualifier, appData2013, featureId, componentId);
                }
            }
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
                case "all":
                    return VisioEdition.All;

                case "x86":
                    return VisioEdition.X86;

                case "x64":
                    return VisioEdition.X64;

                default:
                    Core.OnMessage(VisioErrors.InvalidVisioEdition(sourceLineNumbers, attribValue));
                    return VisioEdition.All;
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
            Unknown,
            Stencil,
            Template,
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
        /// <returns></returns>

        private VisioContentType GetVisioContentType(SourceLineNumberCollection sourceLineNumbers, string fileName)
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

                default:
                    Core.OnMessage(VisioErrors.InvalidFileExtension(sourceLineNumbers, fileName));
                    return VisioContentType.Unknown;
            }
        }

        /// <summary>
        /// Returns the ComponentID for PublishComponent table for given combination of version/content type.
        /// </summary>
        /// <param name="contentType"></param>
        /// <param name="visioVersion"></param>
        /// <returns></returns>

        private string GetVisioComponentId(VisioContentType contentType, VisioVersion visioVersion)
        {
            switch (contentType)
            {
                case VisioContentType.Stencil:
                    return GetVisioStencilComponentId(visioVersion);

                case VisioContentType.Template:
                    return GetVisioTemplateComponentId(visioVersion);

                default:
                    throw new ArgumentOutOfRangeException("contentType");
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="visioVersion"></param>
        /// <returns></returns>

        private string GetVisioTemplateComponentId(VisioVersion visioVersion)
        {
            switch (visioVersion)
            {
                case VisioVersion.Visio2003:
                    return "{CF1F488D-8D6F-499C-A78D-026E1DF38100}";

                case VisioVersion.Visio2007:
                    return "{6D9D8B6F-D0EF-4BC0-8DD4-09DD6CE2B000}";

                case VisioVersion.Visio2010:
                    return "{6D9D8B6F-D0EF-4BC0-8DD4-09DD6CE2B100}";

                case VisioVersion.Visio2013:
                    return "{6D9D8B6F-D0EF-4BC0-8DD4-09DD6CE2B200}";

                default:
                    throw new ArgumentOutOfRangeException("visioVersion");
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="visioVersion"></param>
        /// <returns></returns>

        private static string GetVisioStencilComponentId(VisioVersion visioVersion)
        {
            switch (visioVersion)
            {
                case VisioVersion.Visio2003:
                    return "{CF1F488D-8D6F-499C-A78D-026E1DF38101}";

                case VisioVersion.Visio2007:
                    return "{6D9D8B6F-D0EF-4BC0-8DD4-09DD6CE2B001}";

                case VisioVersion.Visio2010:
                    return "{6D9D8B6F-D0EF-4BC0-8DD4-09DD6CE2B101}";

                case VisioVersion.Visio2013:
                    return "{6D9D8B6F-D0EF-4BC0-8DD4-09DD6CE2B201}";

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

        private uint ParseQuickShapeCount(SourceLineNumberCollection sourceLineNumbers, XmlAttribute attrib)
        {
            string attribValue = Core.GetAttributeBundleVariableValue(sourceLineNumbers, attrib);

            uint result;
            if (uint.TryParse(attribValue, out result))
                return result;

            Core.OnMessage(VisioErrors.InvalidQuickShapesCount(sourceLineNumbers, attribValue));

            return 0;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sourceLineNumbers"></param>
        /// <param name="attrib"></param>
        /// <returns></returns>

        private string ParseAltNames(SourceLineNumberCollection sourceLineNumbers, XmlAttribute attrib)
        {
            return Core.GetAttributeBundleVariableValue(sourceLineNumbers, attrib);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sourceLineNumbers"></param>
        /// <param name="attrib"></param>
        /// <returns></returns>

        private string ParseMenuString(SourceLineNumberCollection sourceLineNumbers, XmlAttribute attrib)
        {
            return Core.GetAttributeBundleVariableValue(sourceLineNumbers, attrib);
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