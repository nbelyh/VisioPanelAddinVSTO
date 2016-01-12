//-------------------------------------------------------------
// GACWrap.cs
//
// This implements managed wrappers to GAC API Interfaces
//-------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Linq;
using System.Text;

namespace PanelAddinWizard
{
    internal class Win32
    {
        [DllImport("fusion.dll")]
        internal static extern int CreateAssemblyEnum(
                        out IAssemblyEnum ppEnum,
                        IntPtr pUnkReserved,
                        IAssemblyName pName,
                        AssemblyCacheFlags flags,
                        IntPtr pvReserved);

        [DllImport("fusion.dll")]
        internal static extern int CreateAssemblyNameObject(
                        out IAssemblyName ppAssemblyNameObj,
                        [MarshalAs(UnmanagedType.LPWStr)]
                String szAssemblyName,
                        CreateAssemblyNameObjectFlags flags,
                        IntPtr pvReserved);
    }

    [Flags]
    internal enum AssemblyCacheFlags
    {
        GAC = 2,
    }

    internal enum CreateAssemblyNameObjectFlags
    {
        CANOF_DEFAULT = 0,
        CANOF_PARSE_DISPLAY_NAME = 1,
    }

    [Flags]
    internal enum AssemblyNameDisplayFlags
    {
        VERSION = 0x01,
        CULTURE = 0x02,
        PUBLIC_KEY_TOKEN = 0x04,
        PROCESSORARCHITECTURE = 0x20,
        RETARGETABLE = 0x80,
        // This enum will change in the future to include
        // more attributes.
        ALL = VERSION
                                                                | CULTURE
                                                                | PUBLIC_KEY_TOKEN
                                                                | PROCESSORARCHITECTURE
                                                                | RETARGETABLE
    }

    [ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("CD193BC0-B4BC-11d2-9833-00C04FC31D2E")]
    internal interface IAssemblyName
    {
        [PreserveSig()]
        int SetProperty(
                        int PropertyId,
                        IntPtr pvProperty,
                        int cbProperty);

        [PreserveSig()]
        int GetProperty(
                        int PropertyId,
                        IntPtr pvProperty,
                        ref int pcbProperty);

        [PreserveSig()]
        int Finalize();

        [PreserveSig()]
        int GetDisplayName(
                        StringBuilder pDisplayName,
                        ref int pccDisplayName,
                        int displayFlags);

        [PreserveSig()]
        int Reserved(ref Guid guid,
                Object obj1,
                Object obj2,
                String string1,
                Int64 llFlags,
                IntPtr pvReserved,
                int cbReserved,
                out IntPtr ppv);

        [PreserveSig()]
        int GetName(
                        ref int pccBuffer,
                        StringBuilder pwzName);

        [PreserveSig()]
        int GetVersion(
                        out int versionHi,
                        out int versionLow);
        [PreserveSig()]
        int IsEqual(
                        IAssemblyName pAsmName,
                        int cmpFlags);

        [PreserveSig()]
        int Clone(out IAssemblyName pAsmName);
    }// IAssemblyName

    [ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("21b8916c-f28e-11d2-a473-00c04f8ef448")]
    internal interface IAssemblyEnum
    {
        [PreserveSig()]
        int GetNextAssembly(
                        IntPtr pvReserved,
                        out IAssemblyName ppName,
                        int flags);
        [PreserveSig()]
        int Reset();
        [PreserveSig()]
        int Clone(out IAssemblyEnum ppEnum);
    }// IAssemblyEnum

    [ComVisible(false)]
    public class AssemblyCacheEnum
    {
        // null means enumerate all the assemblies
        public AssemblyCacheEnum(String assemblyName)
        {
            IAssemblyName fusionName = null;
            int hr = 0;

            if (assemblyName != null)
            {
                hr = Win32.CreateAssemblyNameObject(
                                out fusionName,
                                assemblyName,
                                CreateAssemblyNameObjectFlags.CANOF_PARSE_DISPLAY_NAME,
                                IntPtr.Zero);
            }

            if (hr >= 0)
            {
                hr = Win32.CreateAssemblyEnum(
                                out m_AssemblyEnum,
                                IntPtr.Zero,
                                fusionName,
                                AssemblyCacheFlags.GAC,
                                IntPtr.Zero);
            }

            if (hr < 0)
            {
                Marshal.ThrowExceptionForHR(hr);
            }
        }

        public String GetNextAssembly()
        {
            int hr = 0;
            IAssemblyName fusionName = null;

            if (done)
            {
                return null;
            }

            // Now get next IAssemblyName from m_AssemblyEnum
            hr = m_AssemblyEnum.GetNextAssembly((IntPtr)0, out fusionName, 0);

            if (hr < 0)
            {
                return null;
            }

            if (fusionName != null)
            {
                return GetFullName(fusionName);
            }
            else
            {
                done = true;
                return null;
            }
        }

        private String GetFullName(IAssemblyName fusionAsmName)
        {
            StringBuilder sDisplayName = new StringBuilder(1024);
            int iLen = 1024;

            int hr = fusionAsmName.GetDisplayName(sDisplayName, ref iLen, (int)AssemblyNameDisplayFlags.ALL);
            if (hr < 0)
            {
                return null;
            }

            return sDisplayName.ToString();
        }

        private IAssemblyEnum m_AssemblyEnum = null;
        private bool done;
    }// class AssemblyCacheEnum


    public class GacUtil
    {
        public static int GetInstalledVersion(string name)
        {
            int result = 0;

            IAssemblyName assemblyName;
            result = Win32.CreateAssemblyNameObject(out assemblyName, name, CreateAssemblyNameObjectFlags.CANOF_DEFAULT, IntPtr.Zero);
            if ((result != 0) || (assemblyName == null))
                return 0;

            IAssemblyEnum enumerator;
            result = Win32.CreateAssemblyEnum(out enumerator, IntPtr.Zero, assemblyName, AssemblyCacheFlags.GAC, IntPtr.Zero);
            if ((result != 0) || (enumerator == null))
                return 0;

            while ((enumerator.GetNextAssembly(IntPtr.Zero, out assemblyName, 0) == 0) && (assemblyName != null))
            {
                int versionHi, versionLo;
                assemblyName.GetVersion(out versionHi, out versionLo);
                var version = (versionHi >> 16);
                if (result < version)
                    result = version;
            }

            return result;
        }
    }
}