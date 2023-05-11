using System;
using System.Runtime.InteropServices;

namespace Office_File_Explorer.Helpers
{
    public static class Win32
    {
        // clipboard notifications
        internal const int WM_CLIPBOARDUPDATE = 0x031D;
        internal const int WM_ASKCBFORMATNAME = 0x030C;
        internal const int WM_DRAWCLIPBOARD = 0x0308;
        internal const int WM_CHANGECBCHAIN = 0x030D;
        internal const int WM_DESTROYCLIPBOARD = 0x0307;
        internal const int WM_HSCROLLCLIPBOARD = 0x030E;
        internal const int WM_PAINTCLIPBOARD = 0x309;
        internal const int WM_RENDERALLFORMATS = 0x0306;
        internal const int WM_RENDERFORMAT = 0x0305;
        internal const int WM_SIZECLIPBOARD = 0x030B;
        internal const int WM_VSCROLLCLIPBOARD = 0x030A;

        internal const int CF_METAFILEPICT = 3;
        internal const int CF_ENHMETAFILE = 14;

        // PInvoke declarations
        [DllImport(Strings.user32)]
        internal static extern IntPtr SetClipboardViewer(IntPtr hWndNewViewer);

        [DllImport(Strings.user32)]
        internal static extern IntPtr ChangeClipboardChain(IntPtr hWndRemove, IntPtr hWndNewNext);

        [DllImport(Strings.user32)]
        internal static extern void SendMessage(IntPtr hwnd, uint wMsg, IntPtr wParam, IntPtr lParam);

        [DllImport(Strings.user32, SetLastError = true)]
        internal static extern bool OpenClipboard(IntPtr hWndNewOwner);

        [DllImport(Strings.user32)]
        internal static extern IntPtr GetClipboardData(uint uFormat);

        [DllImport(Strings.user32, SetLastError = true)]
        internal static extern bool CloseClipboard();

        [DllImport(Strings.user32)]
        internal static extern bool IsClipboardFormatAvailable(int wFormat);

        [DllImport(Strings.user32, SetLastError = true)]
        internal static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [DllImport(Strings.user32)]
        internal static extern IntPtr GetClipboardOwner();

        [DllImport(Strings.gdi32)]
        internal static extern IntPtr CopyEnhMetaFile(IntPtr hemfSrc, string lpszFile);

        [DllImport(Strings.gdi32)]
        internal static extern bool DeleteEnhMetaFile(IntPtr hemf);
    }
}
