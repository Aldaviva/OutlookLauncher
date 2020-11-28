#nullable enable

using System;
using System.Runtime.InteropServices;

namespace OutlookLauncher {

    internal static class Win32 {

        [DllImport("user32.dll")]
        internal static extern bool ShowWindow(IntPtr hWnd, ShowWindowCommand nCmdShow);

        internal enum ShowWindowCommand {

            HIDE               = 0,
            NORMAL             = 1,
            SHOW_MINIMIZED     = 2,
            MAXIMIZE           = 3,
            SHOW_NO_ACTIVATE   = 4,
            SHOW               = 5,
            MINIMIZE           = 6,
            SHOW_MIN_NO_ACTIVE = 7,
            SHOW_NA            = 8,
            RESTORE            = 9,
            SHOW_DEFAULT       = 10,
            FORCE_MINIMIZE     = 11

        }

    }

}