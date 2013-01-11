using System;
using System.Runtime.InteropServices;

namespace GoContactSyncMod
{
    internal class WinAPIMethods
    {
        #region API Constants
        public const int HWND_BROADCAST = 0xffff;
        public static readonly int WM_GCSM_SHOWME = RegisterWindowMessage("WM_GCSM_SHOWME");
        public const int VER_NT_WORKSTATION = 0x0000001;

        // Fix for WinXP and older systems, that do not continue with shutdown until all programs have closed
        // FormClosing would hold system shutdown, when it sets the cancel to true
        public const int WM_QUERYENDSESSION = 0x11;

        //Code to find out if workstation is locked
        /*public const int WM_WTSSESSION_CHANGE = 0x02B1;
        public const int WTS_SESSION_LOCK = 0x7;
        public const int WTS_SESSION_UNLOCK = 0x8;

        //Code to find if workstation is resumed
        public const int WM_POWERBROADCAST = 0x0218;
        public const int PBT_APMQUERYSUSPEND = 0x0000;
        public const int PBT_APMQUERYSTANDBY = 0x0001;
        public const int PBT_APMQUERYSUSPENDFAILED = 0x0002;
        public const int PBT_APMQUERYSTANDBYFAILED = 0x0003;
        public const int PBT_APMSUSPEND = 0x0004;
        public const int PBT_APMSTANDBY = 0x0005;
        public const int PBT_APMRESUMECRITICAL = 0x0006;
        public const int PBT_APMRESUMESUSPEND = 0x0007;
        public const int PBT_APMRESUMESTANDBY = 0x0008;
        public const int PBT_APMRESUMEAUTOMATIC = 0x0012;        
        */
        #endregion
 

        [StructLayout(LayoutKind.Sequential)]
        public struct OSVERSIONINFOEX
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

        #region Extern Functions Declaration
        
        [DllImport("user32", SetLastError=true)]
        public static extern bool PostMessage(IntPtr hwnd, int msg, IntPtr wparam, IntPtr lparam);
        [DllImport("user32", SetLastError=true)]
        public static extern int RegisterWindowMessage(string message);
        
        [DllImport("kernel32")]
        public static extern bool GetVersionEx(ref OSVERSIONINFOEX osvi);
        
        //to detect if the user locks or unlocks the workstation
        /*
        [DllImport("wtsapi32.dll")]
        public static extern bool WTSRegisterSessionNotification(IntPtr hWnd, int dwFlags);
        [DllImport("wtsapi32.dll")]
        public static extern bool WTSUnRegisterSessionNotification(IntPtr hWnd);
        */
        #endregion
    }
}
