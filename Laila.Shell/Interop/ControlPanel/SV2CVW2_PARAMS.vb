
Imports System.Runtime.InteropServices

<StructLayout(LayoutKind.Sequential)>
Public Structure SV2CVW2_PARAMS
    Public cbSize As UInt32
    Public psvPrev As IShellView   ' IShellView
    Public pfs As FOLDERSETTINGS   ' LPCFOLDERSETTINGS
    Public psbOwner As IShellBrowser   ' IShellBrowser
    Public prcView As WIN32RECT   ' RECT
    Public pvid As Guid  ' SHELLVIEWID
    Public hwndView As IntPtr  ' HWND
End Structure