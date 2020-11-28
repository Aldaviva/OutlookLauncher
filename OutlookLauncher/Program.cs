#nullable enable

using System;
using System.Diagnostics;
using System.Linq;
using System.Windows.Forms;
using ManagedWinapi.Windows;
using Microsoft.Win32;

namespace OutlookLauncher {

    internal static class Program {

        private static void Main() {
            if (findOutlookWindow() is {} outlookWindow) {
                if (outlookWindow.WindowState == FormWindowState.Minimized) {
                    Win32.ShowWindow(outlookWindow.HWnd, Win32.ShowWindowCommand.RESTORE);
                }

                if (!outlookWindow.TopMost) {
                    outlookWindow.TopMost = true;
                    outlookWindow.TopMost = false;
                }
            } else if (getOutlookExecutableAbsolutePath() is {} executablePath) {
                Process.Start(executablePath)?.Dispose();
            } else {
                MessageBox.Show("Outlook installation not found.", "Failed to start Outlook", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static SystemWindow? findOutlookWindow() => SystemWindow.FilterToplevelWindows(candidate =>
            candidate.ClassName == "rctrl_renwnd32" &&
            candidate.Title.EndsWith(" - Outlook") &&
            "outlook".Equals(candidate.Process.ProcessName, StringComparison.InvariantCultureIgnoreCase)).FirstOrDefault();

        private static string? getOutlookExecutableAbsolutePath() {
            if (Registry.GetValue(@"HKEY_CLASSES_ROOT\stssync\shell\open\command", null, null) is string commandValue) {
                return commandValue.Split(new[] { " /share " }, 2, StringSplitOptions.None)[0];
            } else {
                return null;
            }
        }

    }

}