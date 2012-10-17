using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;

namespace GoContactSyncMod
{
    static class VersionInformation
    {
        public enum OutlookMainVersion
        {
            Outlook2002,
            Outlook2003,
            Outlook2007,
            Outlook2010,
            OutlookUnknownVersion,
            OutlookNoInstance
        }

        public static OutlookMainVersion GetOutlookVersion()
        {
            Microsoft.Office.Interop.Outlook.Application appVersion = new Microsoft.Office.Interop.Outlook.Application();
            switch (appVersion.Version.ToString().Substring(0,2))
            {
                case "10":
                    return OutlookMainVersion.Outlook2002;
                case "11":
                    return OutlookMainVersion.Outlook2003;
                case "12":
                    return OutlookMainVersion.Outlook2007;
                case "14":
                    return OutlookMainVersion.Outlook2010;
                default:
                    {
                        if (appVersion != null)
                        {
                            Marshal.ReleaseComObject(appVersion);
                        }
                        return OutlookMainVersion.OutlookUnknownVersion;
                    }
            }
     
        }

        #region API-Deklarationen

        [StructLayout(LayoutKind.Sequential)]
        private struct OSVERSIONINFOEX
        {
            public int dwOSVersionInfoSize;
            public int dwMajorVersion;
            public int dwMinorVersion;
            public int dwBuildNumber;
            public int dwPlatformId;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 128)]
            public string szCSDVersion;
            public Int16 wServicePackMajor;
            public Int16 wServicePackMinor;
            public Int16 wSuiteMask;
            public Byte wProductType;
            public Byte wReserved;
        }

        [DllImport("kernel32")]
        static extern bool GetVersionEx(ref OSVERSIONINFOEX osvi);

        private const int VER_NT_WORKSTATION = 0x0000001;

        #endregion

        /// <summary>
        /// windows-main-version types
        /// </summary>
        public enum WindowsMainVersion
        {
            WindowsXP,
            WindowsServer2003,
            WindowsXP64,
            Vista,
            WindowsServer2008,
            Windows7,
            WindowsServer2008R2,
            Unknown
        }

        /// <summary>
        /// detect window main version
        /// </summary>
        public static WindowsMainVersion GetWindowsMainVersion()
        {
            OSVERSIONINFOEX osVersionInfo = new OSVERSIONINFOEX();
            osVersionInfo.dwOSVersionInfoSize = Marshal.SizeOf(osVersionInfo);
            if (GetVersionEx(ref osVersionInfo))
            {
                switch (osVersionInfo.dwMajorVersion)
                {
                    case 5:
                        if (Environment.OSVersion.Version.Minor == 1)
                        {
                            return WindowsMainVersion.WindowsXP;
                        }
                        else if (Environment.OSVersion.Version.Minor == 2)
                        {
                            if (osVersionInfo.wProductType == VER_NT_WORKSTATION)
                            {
                                return WindowsMainVersion.WindowsXP64;
                            }
                            else
                            {
                                return WindowsMainVersion.WindowsServer2003;
                            }
                        }
                        else
                        {
                            return WindowsMainVersion.Unknown;
                        }

                    case 6:
                        if (Environment.OSVersion.Version.Minor == 0)
                        {
                            if (osVersionInfo.wProductType == VER_NT_WORKSTATION)
                            {
                                return WindowsMainVersion.Vista;
                            }
                            else
                            {
                                return WindowsMainVersion.WindowsServer2008;
                            }
                        }
                        else if (Environment.OSVersion.Version.Minor == 1)
                        {
                            if (osVersionInfo.wProductType == VER_NT_WORKSTATION)
                            {
                                return WindowsMainVersion.Windows7;
                            }
                            else
                            {
                                return WindowsMainVersion.WindowsServer2008R2;
                            }
                        }
                        return WindowsMainVersion.Unknown;

                    default:
                        return WindowsMainVersion.Unknown;
                }
            }
            return WindowsMainVersion.Unknown;
        }
    }
}
