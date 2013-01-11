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
            Windows8,
            WindowsServer2012,
            Unknown
        }

        /// <summary>
        /// detect window main version
        /// </summary>
        public static WindowsMainVersion GetWindowsMainVersion()
        {
            WinAPIMethods.OSVERSIONINFOEX osVersionInfo = new WinAPIMethods.OSVERSIONINFOEX();
            osVersionInfo.dwOSVersionInfoSize = Marshal.SizeOf(osVersionInfo);
            if (WinAPIMethods.GetVersionEx(ref osVersionInfo))
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
                            if (osVersionInfo.wProductType == WinAPIMethods.VER_NT_WORKSTATION)
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
                            if (osVersionInfo.wProductType == WinAPIMethods.VER_NT_WORKSTATION)
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
                            if (osVersionInfo.wProductType == WinAPIMethods.VER_NT_WORKSTATION)
                            {
                                return WindowsMainVersion.Windows7;
                            }
                            else
                            {
                                return WindowsMainVersion.WindowsServer2008R2;
                            }
                        }
                        else if (Environment.OSVersion.Version.Minor == 2)
                        {
                            if (osVersionInfo.wProductType == WinAPIMethods.VER_NT_WORKSTATION)
                            {
                                return WindowsMainVersion.Windows8;
                            }
                            else
                            {
                                return WindowsMainVersion.WindowsServer2012;
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
