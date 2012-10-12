using System;
using System.Windows.Forms;
using System.Reflection;
using System.Runtime.InteropServices;

namespace GoContactSyncMod
{
    static class ErrorHandler
    {
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
            Seven,
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
                                return WindowsMainVersion.Seven;
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

        private static string OsInfo
        {
            get
            {
                return GetWindowsMainVersion().ToString();
            }
        }

		// TODO: Write a nice error dialog, that maybe supports directly E-Mail sending as bugreport
		public static void Handle(Exception ex)
        {
			Logger.Log(ex.Message, EventType.Error);
			//AppendSyncConsoleText(Logger.GetText());
			Logger.Log("Sync failed.", EventType.Error);

			try
			{
                Program.Instance.ShowBalloonToolTip("Error", ex.Message, ToolTipIcon.Error, 5000);
                /*
				Program.Instance.notifyIcon.BalloonTipTitle = "Error";
				Program.Instance.notifyIcon.BalloonTipText = ex.Message;
				Program.Instance.notifyIcon.BalloonTipIcon = ToolTipIcon.Error;
				Program.Instance.notifyIcon.ShowBalloonTip(5000);
                 */
			}
			catch (Exception)
			{
				// this can fail if form was disposed or not created yet, so catch the exception - balloon is not that important to risk followup error
			}

            string message = "Sorry, an unexpected error occured.\nPlease support us fixing this problem. Go to\nhttps://sourceforge.net/projects/googlesyncmod/ and use the Tracker!\nHint: You can copy this message by pressing CTRL-C in the dialog box.\nPlease check first if error has already been reported.\nProgram Version: {0}\n\nError Details:\n{1}\n\nOS Information:\n{2}";
			message = string.Format(message, AssemblyVersion, ex.ToString(), OsInfo);
            Logger.Log(message, EventType.Debug);
            MessageBox.Show(message, "GO Contact Sync Mod");
        }

		private static string AssemblyVersion
		{
			get
			{
				return Assembly.GetExecutingAssembly().GetName().Version.ToString();
			}
		}
    }
}
