using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using static ManagedShell.Interop.NativeMethods;
using Orientation = System.Windows.Controls.Orientation;
using ActionChangedEventArgs = ManagedShell.Common.SupportingClasses.ActionChangedEventArgs;
using MsgBx = System.Windows.MessageBox;

namespace ManagedShell.Common.Helpers
{
    public static class WindowHelper
    {
        public const string TrayWndClass = "Shell_TrayWnd";

        public static void ShowWindowBottomMost(IntPtr handle)
        {
            SetWindowPos(
                handle,
                (IntPtr)WindowZOrder.HWND_BOTTOM,
                0,
                0,
                0,
                0,
                (int)SetWindowPosFlags.SWP_NOSIZE | (int)SetWindowPosFlags.SWP_NOMOVE | (int)SetWindowPosFlags.SWP_NOACTIVATE/* | SWP_NOZORDER | SWP_NOOWNERZORDER*/);
        }

        public static void ShowWindowTopMost(IntPtr handle)
        {
            SetWindowPos(
                handle,
                (IntPtr)WindowZOrder.HWND_TOPMOST,
                0,
                0,
                0,
                0,
                (int)SetWindowPosFlags.SWP_NOSIZE | (int)SetWindowPosFlags.SWP_NOMOVE | (int)SetWindowPosFlags.SWP_SHOWWINDOW/* | (int)SetWindowPosFlags.SWP_NOACTIVATE | SWP_NOZORDER | SWP_NOOWNERZORDER*/);
        }

        public static void ShowWindowDesktop(IntPtr hwnd)
        {
            IntPtr desktopHwnd = GetLowestDesktopParentHwnd();

            if (desktopHwnd != IntPtr.Zero)
            {
                IntPtr nextHwnd = GetWindow(desktopHwnd, GetWindow_Cmd.GW_HWNDPREV);
                SetWindowPos(
                    hwnd,
                    nextHwnd,
                    0,
                    0,
                    0,
                    0,
                    (int)SetWindowPosFlags.SWP_NOSIZE | (int)SetWindowPosFlags.SWP_NOMOVE | (int)SetWindowPosFlags.SWP_NOACTIVATE);
            }
            else
            {
                ShowWindowBottomMost(hwnd);
            }
        }

        public static IntPtr GetLowestDesktopParentHwnd()
        {
            IntPtr progmanHwnd = FindWindow("Progman", "Program Manager");
            IntPtr desktopHwnd = FindWindowEx(progmanHwnd, IntPtr.Zero, "SHELLDLL_DefView", null);

            if (desktopHwnd == IntPtr.Zero)
            {
                IntPtr workerHwnd = IntPtr.Zero;
                IntPtr shellIconsHwnd;
                do
                {
                    workerHwnd = FindWindowEx(IntPtr.Zero, workerHwnd, "WorkerW", null);
                    shellIconsHwnd = FindWindowEx(workerHwnd, IntPtr.Zero, "SHELLDLL_DefView", null);
                } while (shellIconsHwnd == IntPtr.Zero && workerHwnd != IntPtr.Zero);

                desktopHwnd = workerHwnd;
            }
            else
            {
                desktopHwnd = progmanHwnd;
            }

            return desktopHwnd;
        }

        public static IntPtr GetLowestDesktopChildHwnd()
        {
            IntPtr progmanHwnd = FindWindow("Progman", "Program Manager");
            IntPtr desktopHwnd = FindWindowEx(progmanHwnd, IntPtr.Zero, "SHELLDLL_DefView", null);

            if (desktopHwnd == IntPtr.Zero)
            {
                IntPtr workerHwnd = IntPtr.Zero;
                IntPtr shellIconsHwnd;
                do
                {
                    workerHwnd = FindWindowEx(IntPtr.Zero, workerHwnd, "WorkerW", null);
                    shellIconsHwnd = FindWindowEx(workerHwnd, IntPtr.Zero, "SHELLDLL_DefView", null);
                } while (shellIconsHwnd == IntPtr.Zero && workerHwnd != IntPtr.Zero);

                desktopHwnd = shellIconsHwnd;
            }

            return desktopHwnd;
        }
        
        public static void HideWindowFromTasks(IntPtr hWnd)
        {
            SetWindowLong(hWnd, GWL_EXSTYLE, GetWindowLong(hWnd, GWL_EXSTYLE) | (int)ExtendedWindowStyles.WS_EX_TOOLWINDOW);

            ExcludeWindowFromPeek(hWnd);
        }

        public static void ExcludeWindowFromPeek(IntPtr hWnd)
        {
            int status = (int)DWMNCRENDERINGPOLICY.DWMNCRP_ENABLED;
            DwmSetWindowAttribute(hWnd,
                DWMWINDOWATTRIBUTE.DWMWA_EXCLUDED_FROM_PEEK,
                ref status,
                sizeof(int));
        }

        public static void PeekWindow(bool show, IntPtr targetHwnd, IntPtr callingHwnd)
        {
            uint enable = 0;
            if (show) enable = 1;

            if (EnvironmentHelper.IsWindows81OrBetter)
            {
                DwmActivateLivePreview(enable, targetHwnd, callingHwnd, AeroPeekType.Window, IntPtr.Zero);
            }
            else
            {
                DwmActivateLivePreview(enable, targetHwnd, callingHwnd, AeroPeekType.Window);
            }
        }

        public static void SetWindowBlur(IntPtr hWnd, bool enable)
        {
            if (EnvironmentHelper.IsWindows10OrBetter)
            {
                // https://github.com/riverar/sample-win32-acrylicblur
                // License: MIT
                var accent = new AccentPolicy();
                var accentStructSize = Marshal.SizeOf(accent);
                if (enable)
                {
                    if (EnvironmentHelper.IsWindows10RS4OrBetter)
                    {
                        accent.AccentState = AccentState.ACCENT_ENABLE_ACRYLICBLURBEHIND;
                        accent.GradientColor = (0 << 24) | (0xFFFFFF /* BGR */ & 0xFFFFFF);
                    }
                    else
                    {
                        accent.AccentState = AccentState.ACCENT_ENABLE_BLURBEHIND;
                    }
                }
                else
                {
                    accent.AccentState = AccentState.ACCENT_DISABLED;
                }

                var accentPtr = Marshal.AllocHGlobal(accentStructSize);
                Marshal.StructureToPtr(accent, accentPtr, false);

                var data = new WindowCompositionAttributeData();
                data.Attribute = WindowCompositionAttribute.WCA_ACCENT_POLICY;
                data.SizeOfData = accentStructSize;
                data.Data = accentPtr;

                SetWindowCompositionAttribute(hWnd, ref data);

                Marshal.FreeHGlobal(accentPtr);
            }
        }

        public static bool SetDarkModePreference(PreferredAppMode mode)
        {
            if (EnvironmentHelper.IsWindows10DarkModeSupported)
            {
                return SetPreferredAppMode(mode);
            }

            return false;
        }

        public static IntPtr FindWindowsTray(IntPtr hwndIgnore)
        {
            IntPtr taskbarHwnd = FindWindow(TrayWndClass, "");

            if (hwndIgnore != IntPtr.Zero)
            {
                while (taskbarHwnd == hwndIgnore)
                {
                    taskbarHwnd = FindWindowEx(IntPtr.Zero, taskbarHwnd, TrayWndClass, "");
                }
            }

            return taskbarHwnd;
        }

        static Action _undoArrangeDesktopWindows = null;
        public static Action UndoArrangeDesktopWindows
        {
            get => _undoArrangeDesktopWindows;
            private set
            {
                _undoArrangeDesktopWindows = value;
                ArrangeDesktopWindowsUndoChanged?.Invoke(null, new ActionChangedEventArgs(_undoArrangeDesktopWindows));
            }
        }

        public static void CascadeDesktopWindows()
            => ArrangeDesktopWindows((cKids, lpKids) => CascadeWindows(IntPtr.Zero, TileHowFlags.MDITILE_ZORDER, IntPtr.Zero, cKids, lpKids));

        public static void StackDesktopWindows(Orientation orientation)
            => ArrangeDesktopWindows((cKids, lpKids) => TileWindows(IntPtr.Zero, (orientation == Orientation.Vertical) ? TileHowFlags.MDITILE_VERTICAL : TileHowFlags.MDITILE_HORIZONTAL, IntPtr.Zero, cKids, lpKids));

        static void ArrangeDesktopWindows(Action<int, IntPtr[]> arrange)
        {
            Dictionary<IntPtr, Rect> allPrevBounds = new Dictionary<IntPtr, Rect>();
            Dictionary<IntPtr, WindowShowStyle> allPrevShowCmd = new Dictionary<IntPtr, WindowShowStyle>();
            EnumDesktopWindows(IntPtr.Zero, (hwnd, lParam) =>
            {
                int style = GetWindowLong(hwnd, GWL_EXSTYLE);
                if (
                    IsWindow(hwnd)
                    && IsWindowVisible(hwnd)
                    && GetWindowRect(hwnd, out Rect bounds)
                    && TryGetWindowShowStyle(hwnd, out WindowShowStyle showStyle)
                )
                {
                    allPrevBounds[hwnd] = bounds;
                    allPrevShowCmd[hwnd] = showStyle;
                }
                return true;
            }, 0);

            IntPtr[] hwnds = allPrevBounds.Keys.ToArray();
            
            if (hwnds.Length <= 0)
                return;

            arrange(0, null);
            
            for (int winIndex = 0; winIndex < hwnds.Length; winIndex++)
            {
                IntPtr hwnd = hwnds[winIndex];
                bool remove = true;

                if (GetWindowRect(hwnd, out Rect newBounds)
                    && TryGetWindowShowStyle(hwnd, out WindowShowStyle newShowStyle)
                    )
                {
                    Rect oldBounds = allPrevBounds[hwnd];
                    WindowShowStyle oldShowStyle = allPrevShowCmd[hwnd];
                    if (
                    
                           (newBounds.Left != oldBounds.Left)
                        || (newBounds.Top != oldBounds.Top)
                        || (newBounds.Right != oldBounds.Right)
                        || (newBounds.Bottom != oldBounds.Bottom)
                        || (newShowStyle != oldShowStyle)
                        )
                    {
                        remove = false;
                    }
                }

                if (remove)
                {
                    allPrevBounds.Remove(hwnd);
                    allPrevShowCmd.Remove(hwnd);
                }
            }
            
            hwnds = allPrevBounds.Keys.ToArray();

            UndoArrangeDesktopWindows = () =>
            {
                for (int winIndex = 0; winIndex < hwnds.Length; winIndex++)
                {
                    IntPtr hwnd = hwnds[winIndex];
                    
                    Rect prevBounds = allPrevBounds[hwnd];
                    WindowShowStyle prevShowStyle = allPrevShowCmd[hwnd];
                    
                    ShowWindow(hwnd, prevShowStyle);
                    SetWindowPos(hwnd, IntPtr.Zero, prevBounds.Left, prevBounds.Top, prevBounds.Width, prevBounds.Height, (int)(SetWindowPosFlags.SWP_NOZORDER));
                }
                UndoArrangeDesktopWindows = null;
            };
        }

        static bool TryGetWindowShowStyle(IntPtr hwnd, out WindowShowStyle showStyle)
        {
            WINDOWPLACEMENT placement = new WINDOWPLACEMENT();
            if (GetWindowPlacement(hwnd, ref placement))
            {
                showStyle = placement.showCmd;
                return true;
            }

            showStyle = (WindowShowStyle)(13379001);
            return false;
        }

        public static void ToggleDesktop()
        {
            Process.Start(new ProcessStartInfo("shell:::{3080F90D-D7AD-11D9-BD98-0000947B0257}")
            {
                UseShellExecute = true
            });
        }

        public static event EventHandler<ActionChangedEventArgs> ArrangeDesktopWindowsUndoChanged;
    }
}
