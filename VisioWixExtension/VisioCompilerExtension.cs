using System;
using System.Collections;
using System.Reflection;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Schema;
using Microsoft.Tools.WindowsInstallerXml;
using System.IO;
using System.Globalization;

namespace VisioWixExtension
{
    internal class VisioCompilerExtension : CompilerExtension
    {
        private XmlSchema schema;
        private Hashtable components;

        /// <summary>
        /// Instantiate a new DifxAppCompiler.
        /// </summary>
        public VisioCompilerExtension()
        {
            this.components = new Hashtable();
        }

        /// <summary>
        /// Gets the schema for this extension.
        /// </summary>
        /// <value>Schema for this extension.</value>
        public override XmlSchema Schema
        {
            get
            {
                if (null == this.schema)
                {
                    this.schema = LoadXmlSchemaHelper(Assembly.GetExecutingAssembly(), "VisioWixExtension.Xsd.visio.xsd");
                }

                return this.schema;
            }
        }

        private const string WixNameSpace = "http://schemas.microsoft.com/wix/2006/wi";

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

                    switch (element.LocalName)
                    {
                        case "Publish":
                            this.ParseVisioElement(element, fileId, componentId);
                            break;
                        default:
                            this.Core.UnexpectedElement(parentElement, element);
                            break;
                    }
                    break;
                default:
                    this.Core.UnexpectedElement(parentElement, element);
                    break;
            }
        }

        [Flags]
        private enum VisioVersion
        {
            None = 0,
            Visio2003 = 1,
            Visio2007 = 2,
            Visio2010 = 4,
            Visio2013 = 8,
        };

        /// <summary>
        /// Parses a Driver element.
        /// </summary>
        /// <param name="node">Element to parse.</param>
        /// <param name="componentId">Identifier for parent component.</param>
        private void ParseVisioElement(XmlNode node, string contextFileId, string contextComponentId)
        {
            SourceLineNumberCollection sourceLineNumbers = Preprocessor.GetSourceLineNumbers(node);
            int attributes = 0;
            int sequence = CompilerCore.IntegerNotSet;

            string component = null;
            string feature = null;

            VisioVersion visioVersion = GetDefaultVisioVersion(contextFileId);
            VisioContentType visioContentType = GetVisioContentType(contextFileId);

            string visioEdition = "All";
            string visioLanguage = "All";
            string menuPath = "";
            string altNames = "";
            string quickShapesCount = "0";

            foreach (XmlAttribute attrib in node.Attributes)
            {
                switch (attrib.LocalName)
                {
                    case "Feature":
                        feature = this.Core.GetAttributeIdentifierValue(sourceLineNumbers, attrib);
                        this.Core.CreateWixSimpleReferenceRow(sourceLineNumbers, "Feature", feature);
                        break;

                    case "Component":
                        component = this.Core.GetAttributeIdentifierValue(sourceLineNumbers, attrib);
                        this.Core.CreateWixSimpleReferenceRow(sourceLineNumbers, "Component", component);
                        break;

                    case "Visio2003":
                        if (YesNoType.Yes == this.Core.GetAttributeYesNoValue(sourceLineNumbers, attrib))
                            visioVersion |= VisioVersion.Visio2003;
                        break;

                    case "Visio2007":
                        if (YesNoType.Yes == this.Core.GetAttributeYesNoValue(sourceLineNumbers, attrib))
                            visioVersion |= VisioVersion.Visio2007;
                        break;

                    case "Visio2010":
                        if (YesNoType.Yes == this.Core.GetAttributeYesNoValue(sourceLineNumbers, attrib))
                            visioVersion |= VisioVersion.Visio2010;
                        break;

                    case "Visio2013":
                        if (YesNoType.Yes == this.Core.GetAttributeYesNoValue(sourceLineNumbers, attrib))
                            visioVersion |= VisioVersion.Visio2013;
                        break;

                    case "VisioEdition":
                        visioEdition = this.Core.GetAttributeBundleVariableValue(sourceLineNumbers, attrib);
                        break;

                    case "Language":
                        visioLanguage = this.Core.GetAttributeBundleVariableValue(sourceLineNumbers, attrib);
                        break;

                    case "MenuPath":
                        menuPath = this.Core.GetAttributeBundleVariableValue(sourceLineNumbers, attrib);
                        break;

                    case "AltNames":
                        altNames = this.Core.GetAttributeBundleVariableValue(sourceLineNumbers, attrib);
                        break;

                    case "QuickShapesCount":
                        quickShapesCount = this.Core.GetAttributeBundleVariableValue(sourceLineNumbers, attrib);
                        break;
                }
            }

            if (!this.Core.EncounteredError)
            {
                int languageCode = GetVisioLanguageCode(visioLanguage);

                string featureId = (null == feature) ? Guid.Empty.ToString("B") : feature;
                string componentId = (null == component) ? contextComponentId : component;
                string appData = string.Format(@"{0}|{1}|{2}|{3}", menuPath, altNames, quickShapesCount,
                                               GetVisioEditionCode(visioEdition));
                string qualifier = string.Format(@"{0}\{1}", languageCode, contextFileId);

                if ((visioVersion & VisioVersion.Visio2003) == VisioVersion.Visio2003)
                {
                    if (languageCode == 1)
                    {
                        foreach (var lcid in GetAllVisioLanguageCodes())
                        {
                            qualifier = string.Format(@"{0}\{1}", lcid, contextFileId);

                            GenerateRow(sourceLineNumbers, GetVisioComponentId(visioContentType, VisioVersion.Visio2007),
                                        qualifier, appData, featureId, componentId);
                        }
                    }
                    else
                    {
                        GenerateRow(sourceLineNumbers, GetVisioComponentId(visioContentType, VisioVersion.Visio2003),
                                    qualifier, appData, featureId, componentId);
                    }
                }

                if ((visioVersion & VisioVersion.Visio2007) == VisioVersion.Visio2007)
                    GenerateRow(sourceLineNumbers, GetVisioComponentId(visioContentType, VisioVersion.Visio2007),
                                qualifier, appData, featureId, componentId);

                if ((visioVersion & VisioVersion.Visio2010) == VisioVersion.Visio2010)
                    GenerateRow(sourceLineNumbers, GetVisioComponentId(visioContentType, VisioVersion.Visio2010),
                                qualifier, appData, featureId, componentId);

                if ((visioVersion & VisioVersion.Visio2013) == VisioVersion.Visio2013)
                    GenerateRow(sourceLineNumbers, GetVisioComponentId(visioContentType, VisioVersion.Visio2013),
                                qualifier, appData, featureId, componentId);
            }
        }

        private void GenerateRow(SourceLineNumberCollection sourceLineNumbers, string visioComponentId, string qualifier,
                                 string appData, string featureId, string componentId)
        {
            Row row = this.Core.CreateRow(sourceLineNumbers, "PublishComponent");

            row[0] = visioComponentId;
            row[1] = qualifier;
            row[2] = componentId;
            row[3] = appData;
            row[4] = featureId;
        }

        private string GetVisioEditionCode(string visioEdition)
        {
            switch (visioEdition)
            {
                case "All":
                    return "-1";

                case "32bit":
                    return "32";

                case "64bit":
                    return "64";

                default:
                    return "";
            }
        }

        private enum VisioContentType
        {
            Unknown,
            Stencil,
            Template,
        };

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
                    return VisioVersion.Visio2013;

                default:
                    return VisioVersion.None;
            }
        }

        private VisioContentType GetVisioContentType(string fileName)
        {
            switch (Path.GetExtension(fileName))
            {
                case ".vss":
                case ".vsx":
                case ".vssx":
                    return VisioContentType.Stencil;

                case ".vst":
                case ".vtx":
                case ".vstx":
                    return VisioContentType.Template;

                default:
                    return VisioContentType.Unknown;
            }
        }

        private string GetVisioComponentId(VisioContentType contentType, VisioVersion visioVersion)
        {
            switch (contentType)
            {
                case VisioContentType.Stencil:
                    return GetVisioStencilComponentId(visioVersion);

                case VisioContentType.Template:
                    return GetVisioTemplateComponentId(visioVersion);

                default:
                    return "";
            }
        }

        private string GetVisioStencilComponentId(VisioVersion visioVersion)
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
                    return "";
            }
        }

        private string GetVisioTemplateComponentId(VisioVersion visioVersion)
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
                    return "";
            }
        }

        private int GetVisioLanguageCode(string visioLanguage)
        {
            if (visioLanguage == "All")
                return 1;

            return CultureInfo.GetCultureInfo(visioLanguage).LCID;
        }

        private int[] GetAllVisioLanguageCodes()
        {
            return new int[]
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