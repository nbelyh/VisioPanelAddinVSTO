using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;

namespace VisioWixExtension
{
    /// <summary>
    /// Type of Visio content to register: template/stencil/etc
    /// </summary>
    enum VisioContentType
    {
        Template = 0,
        Stencil = 1,
        Help = 2,
        Addon = 3,

        Unknown = -1
    };

    /// <summary>
    /// Supported Visio versions (2003/2007/2010/etc)
    /// </summary>
    [Flags]
    enum VisioVersion
    {
        Visio2003 = 1,
        Visio2007 = 2,
        Visio2010 = 4,
        Visio2013 = 8,
        Visio2016 = 16,

        Default = 0
    };

    /// <summary>
    /// Supported Visio editions (All/32bit/64bit)
    /// </summary>
    [Flags]
    enum VisioEdition
    {
        X86 = 1,
        X64 = 2,

        Default = 0,
    };

    /// <summary>
    /// Supported Visio Addon Attributes (see Xsd for more information)
    /// </summary>
    [Flags]
    enum AddonAttrs
    {
        PerformsActions = 1,
        HasAboutBox = 2,
        ProvidesHelp = 4,
        DisplayWaitCursor = 8,
        HideInUI = 16,
    };

    /// <summary>
    /// Supported Visio Addon Enabling Policy (see Xsd for more information)
    /// </summary>
    [Flags]
    internal enum EnablingPolicy
    {
        AlwaysEnabled = 0xffff,
        DynamicallyEnabled = 0x0,
        StaticallyEnabled = 0x1,
        StaticallyThenDynamicallyEnabled = 0x8001
    }

    /// <summary>
    /// Supported Visio Addon Enabling Static Context (see Xsd for more information)
    /// </summary>
    [Flags]
    internal enum StaticEnableConditions
    {
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
    /// All information needed to generate PublishComponent rows for a single file
    /// </summary>
    class VisioPublishInfo
    {
        public VisioPublishInfo(string fileName)
        {
            VisioVersion = VisioVersion.Default;
            VisioEdition = VisioEdition.Default;
            AddonAttrs = AddonAttrs.PerformsActions;

            MenuPath = Path.GetFileName(fileName);
            AltNames = "";

            VisioLanguage = 1;

            QuickShapesCount = 0;
            FeaturedTemplate = true;

            LocalizedName = Path.GetFileName(fileName);
            UniversalName = Path.GetFileName(fileName);

            EnablingPolicy = EnablingPolicy.AlwaysEnabled;
            StaticEnableConditions = StaticEnableConditions.Document;

            Ordinal = 1;
        }

        public VisioVersion VisioVersion;
        public VisioEdition VisioEdition;

        public AddonAttrs AddonAttrs;

        public string MenuPath;
        public string AltNames;

        public int VisioLanguage;

        public uint QuickShapesCount;
        public bool FeaturedTemplate;

        public string LocalizedName;
        public string UniversalName;

        public uint Ordinal;

        public EnablingPolicy EnablingPolicy;
        public StaticEnableConditions StaticEnableConditions;
        public bool InvokeOnStartup;

        /// <summary>
        /// Returns default set of Visio versions to install. By Default, it is 2007/2010/2013/2016.
        /// For new file types (2013 and above) the default is to publish for Visio 2013 and above only.
        /// </summary>
        /// <param name="fileName">File name to analyze</param>
        /// <returns>Set of Visio version to publish for by default</returns>
        static public VisioVersion GetDefaultVisioVersion(string fileName)
        {
            switch (Path.GetExtension(fileName))
            {
                case ".vssx":
                case ".vstx":
                case ".vstm":
                    return VisioVersion.Visio2013 | VisioVersion.Visio2016;

                default:
                    return VisioVersion.Visio2007 | VisioVersion.Visio2010 | VisioVersion.Visio2013 | VisioVersion.Visio2016;
            }
        }

        /// <summary>
        /// Returns the ComponentID for PublishComponent table for given combination of version/content type.
        /// Visio Component ids are built using the following rule:
        /// {XXXXXX-XXXX-XXX-XXXXXY} - here X are fixed "base" digits, 
        /// whereas the last one (Y) is determined by content type.
        /// </summary>
        /// <param name="contentType"></param>
        /// <param name="visioVersion"></param>
        /// <returns></returns>

        private static string GetVisioComponentId(VisioContentType contentType, VisioVersion visioVersion)
        {
            var pattern = GetVisioBaseComponentId(visioVersion);

            return pattern.Replace("X", Convert.ToInt32(contentType).ToString(CultureInfo.InvariantCulture));
        }

        /// <summary>
        /// Returns base Visio component id based on Visio version
        /// These values were taken from MSDN documentation on PublishComponent table.
        /// </summary>
        /// <param name="visioVersion"></param>
        /// <returns></returns>

        private static string GetVisioBaseComponentId(VisioVersion visioVersion)
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

                case VisioVersion.Visio2016:
                    return "{6D9D8B6F-D0EF-4BC0-8DD4-09DD6CE2B30X}";

                default:
                    throw new ArgumentOutOfRangeException("visioVersion");
            }
        }

        /// <summary>
        /// Returns "edition code" used to build "AppData" field in the PublishComponent table.
        /// </summary>
        /// <param name="visioEdition"></param>
        /// <returns></returns>
        public static string GetVisioEditionCode(VisioEdition visioEdition)
        {
            switch (visioEdition)
            {
                case (VisioEdition.X86|VisioEdition.X64):
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
        /// Builds a "Qualifier" cell of teh "PublishComponent" table out of it's parts.
        /// </summary>
        /// <param name="visioContentType"></param>
        /// <param name="lcid"></param>
        /// <param name="fileName"></param>
        /// <param name="ordinal"></param>
        /// <returns></returns>
        public static string MakeQualifier(VisioContentType visioContentType, int lcid, string fileName, uint ordinal)
        {
            return (visioContentType == VisioContentType.Addon)
                ? String.Format(@"{0}\{1}\{2}", lcid, ordinal - 1, fileName)
                : String.Format(@"{0}\{1}", lcid, fileName);
        }

        /// <summary>
        /// List of all languages supported by Visio 2003. We need to add those explicitly for Visio 2003
        /// </summary>
        /// <returns></returns>
        public static IEnumerable<int> GetAllVisioLanguageCodes()
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


        /// <summary>
        /// Information sufficient to build a single row in PublishComponent table
        /// </summary>
        public class RowInfo
        {
            public string VisioComponentId;
            public string Qualifier;
            public string AppData;
        }


        /// <summary>
        /// Returns the list of PublishComponent rows
        /// </summary>
        /// <param name="visioContentType"></param>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public IEnumerable<RowInfo> GenerateRowInfos(VisioContentType visioContentType,  string fileName)
        {
            var result = new List<RowInfo>();

            var editionCode = GetVisioEditionCode(VisioEdition);

            var enablingPolicyCode = (int) (EnablingPolicy) | (int) (StaticEnableConditions);
            var addonAttrsCode = (int) (AddonAttrs);
            var invokeOnStartupCode = (InvokeOnStartup ? 1 : 0);

            if ((VisioVersion & VisioVersion.Visio2003) == VisioVersion.Visio2003)
            {
                string appData2003 = "";

                switch (visioContentType)
                {
                    case VisioContentType.Stencil:
                    case VisioContentType.Template:
                        appData2003 = String.Format(@"{0}|{1}", MenuPath, AltNames);
                        break;

                    case VisioContentType.Help:
                        appData2003 = null;
                        break;

                    case VisioContentType.Addon:
                        appData2003 = String.Format(@"{0}|{1}|{2}|{3}|{4}|{5}|{6}",
                                                    MenuPath, LocalizedName, UniversalName,
                                                    Ordinal, addonAttrsCode, enablingPolicyCode, invokeOnStartupCode);
                        break;
                }

                if (VisioLanguage == 1)
                {
                    foreach (var lcid in GetAllVisioLanguageCodes())
                    {
                        result.Add(
                            new RowInfo
                                {
                                    VisioComponentId = GetVisioComponentId(visioContentType, VisioVersion.Visio2003),
                                    Qualifier = MakeQualifier(visioContentType, lcid, fileName, Ordinal),
                                    AppData = appData2003
                                });
                    }
                }
                else
                {
                    result.Add(
                        new RowInfo
                        {
                            VisioComponentId = GetVisioComponentId(visioContentType, VisioVersion.Visio2003), 
                            Qualifier = MakeQualifier(visioContentType, VisioLanguage, fileName, Ordinal), 
                            AppData = appData2003
                        });
                }
            }

            var qualifier = MakeQualifier(visioContentType, VisioLanguage, fileName, Ordinal);

            if ((VisioVersion & VisioVersion.Visio2007) == VisioVersion.Visio2007)
            {
                string appData2007 = "";

                switch (visioContentType)
                {
                    case VisioContentType.Stencil:
                        appData2007 = String.Format(@"{0}|{1}", MenuPath, AltNames);
                        break;

                    case VisioContentType.Template:
                        appData2007 = String.Format(@"{0}|{1}|{2}", MenuPath, AltNames,
                                                    FeaturedTemplate ? 1 : 0);
                        break;

                    case VisioContentType.Help:
                        appData2007 = "-1";
                        break;

                    case VisioContentType.Addon:
                        appData2007 = String.Format(@"{0}|{1}|{2}|{3}|{4}|{5}|{6}",
                                                    MenuPath, LocalizedName, UniversalName,
                                                    Ordinal, addonAttrsCode, enablingPolicyCode, invokeOnStartupCode);
                        break;
                }

                result.Add(
                    new RowInfo
                    {
                        VisioComponentId = GetVisioComponentId(visioContentType, VisioVersion.Visio2007),
                        Qualifier = qualifier,
                        AppData = appData2007
                    });
            }

            if ((VisioVersion & VisioVersion.Visio2010) == VisioVersion.Visio2010)
            {
                string appData2010 = "";
                switch (visioContentType)
                {
                    case VisioContentType.Stencil:
                        appData2010 = String.Format(@"{0}|{1}|{2}|{3}", MenuPath, AltNames,
                                                    QuickShapesCount, editionCode);
                        break;

                    case VisioContentType.Template:
                        appData2010 = String.Format(@"{0}|{1}|{2}|{3}", MenuPath, AltNames, 1,
                                                    editionCode);
                        break;

                    case VisioContentType.Help:
                        appData2010 = "-1";
                        break;

                    case VisioContentType.Addon:
                        appData2010 = String.Format(@"{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}",
                                                    MenuPath, LocalizedName, UniversalName,
                                                    Ordinal, addonAttrsCode, enablingPolicyCode, invokeOnStartupCode,
                                                    editionCode);
                        break;
                }

                result.Add(
                    new RowInfo
                    {
                        VisioComponentId = GetVisioComponentId(visioContentType, VisioVersion.Visio2010),
                        Qualifier = qualifier,
                        AppData = appData2010
                    });
            }

            if ((VisioVersion & VisioVersion.Visio2013) == VisioVersion.Visio2013)
            {
                string appData2013 = "";

                switch (visioContentType)
                {
                    case VisioContentType.Stencil:
                        appData2013 = String.Format(@"{0}|{1}|{2}|{3}", MenuPath, AltNames,
                                                    QuickShapesCount, editionCode);
                        break;

                    case VisioContentType.Template:
                        appData2013 = String.Format(@"{0}|{1}|{2}|{3}", MenuPath, AltNames, 1,
                                                    editionCode);
                        break;

                    case VisioContentType.Help:
                        appData2013 = "-1";
                        break;

                    case VisioContentType.Addon:
                        appData2013 = String.Format(@"{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}",
                                                    MenuPath, LocalizedName, UniversalName,
                                                    Ordinal, addonAttrsCode, enablingPolicyCode, invokeOnStartupCode,
                                                    editionCode);
                        break;
                }

                result.Add(
                    new RowInfo
                    {
                        VisioComponentId = GetVisioComponentId(visioContentType, VisioVersion.Visio2013),
                        Qualifier = qualifier,
                        AppData = appData2013
                    });
            }

            if ((VisioVersion & VisioVersion.Visio2016) == VisioVersion.Visio2016)
            {
                string appData2016 = "";

                switch (visioContentType)
                {
                    case VisioContentType.Stencil:
                        appData2016 = String.Format(@"{0}|{1}|{2}|{3}", MenuPath, AltNames,
                                                    QuickShapesCount, editionCode);
                        break;

                    case VisioContentType.Template:
                        appData2016 = String.Format(@"{0}|{1}|{2}|{3}", MenuPath, AltNames, 1,
                                                    editionCode);
                        break;

                    case VisioContentType.Help:
                        appData2016 = "-1";
                        break;

                    case VisioContentType.Addon:
                        appData2016 = String.Format(@"{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}",
                                                    MenuPath, LocalizedName, UniversalName,
                                                    Ordinal, addonAttrsCode, enablingPolicyCode, invokeOnStartupCode,
                                                    editionCode);
                        break;
                }

                result.Add(
                    new RowInfo
                    {
                        VisioComponentId = GetVisioComponentId(visioContentType, VisioVersion.Visio2016),
                        Qualifier = qualifier,
                        AppData = appData2016
                    });
            }

            return result;
        }
    }
}
