Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports Laila.Shell.Interop.Windows

Namespace Interop.Folders
    <ComImport, Guid("000214e3-0000-0000-c000-000000000046")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IShellView
        <PreserveSig>
        Function GetWindow(ByRef hwnd As IntPtr) As Integer
        <PreserveSig>
        Function ContextSensitiveHelp(fEnterMode As Integer) As Integer
        <PreserveSig>
        Function TranslateAccelerator(ByRef lpmsg As Message) As Long

        Sub EnableModeless(fEnable As Boolean)
        Sub UIActivate(<MarshalAs(UnmanagedType.U4)> uState As UInt32)
        Sub Refresh()
        <PreserveSig>
        Function CreateViewWindow(<MarshalAs(UnmanagedType.Interface)> psvPrevious As IShellView, ByRef pfs As FOLDERSETTINGS, <MarshalAs(UnmanagedType.Interface)> psb As IShellBrowser, ByRef prcVie As WIN32RECT, ByRef phWnd As IntPtr) As Integer
        Sub DestroyViewWindow()
        Sub GetCurrentInfo(ByRef lpfs As FOLDERSETTINGS)
        Sub AddPropertySheetPages(<MarshalAs(UnmanagedType.U4)> dwReserved As UInt32, ByRef lpfn As IntPtr, lparam As IntPtr)
        Sub SaveViewState()
        Sub SelectItem(pidlItem As IntPtr, <MarshalAs(UnmanagedType.U4)> uFlags As UInt32)
        <PreserveSig>
        Function GetItemObject(uItem As UInt32, ByRef riid As Guid, ByRef ppv As IntPtr) As Integer
    End Interface
End Namespace
