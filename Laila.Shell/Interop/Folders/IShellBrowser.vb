Imports System.Runtime.InteropServices

Namespace Interop.Folders
    <ComImport, Guid("000214e2-0000-0000-c000-000000000046")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IShellBrowser
        <PreserveSig>
        Function GetWindow(ByRef hwnd As IntPtr) As Integer
        <PreserveSig>
        Function ContextSensitiveHelp(fEnterMode As Integer) As Integer
        <PreserveSig>
        Function InsertMenusSB(hmenuShared As IntPtr, lpMenuWidths As IntPtr) As Integer
        <PreserveSig>
        Function SetMenuSB(hmenuShared As IntPtr, holemenuRes As IntPtr, hwndActiveObject As IntPtr) As Integer
        <PreserveSig>
        Function RemoveMenusSB(hmenuShared As IntPtr) As Integer
        <PreserveSig>
        Function SetStatusTextSB(pszStatusText As IntPtr) As Integer
        <PreserveSig>
        Function EnableModelessSB(fEnable As Boolean) As Integer
        <PreserveSig>
        Function TranslateAcceleratorSB(pmsg As IntPtr, wID As Short) As Integer
        <PreserveSig>
        Function BrowseObject(pidl As IntPtr, <MarshalAs(UnmanagedType.U4)> wFlags As UInt32)
        <PreserveSig>
        Function GetViewStateStream(grfMode As UInt32, ppStrm As IntPtr) As Integer
        <PreserveSig>
        Function GetControlWindow(id As UInt32, ByRef phwnd As IntPtr) As Integer
        <PreserveSig>
        Function SendControlMsg(id As UInt32, uMsg As UInt32, wParam As UInt32, lParam As UInt32, pret As IntPtr) As Integer
        <PreserveSig>
        Function QueryActiveShellView(<MarshalAs(UnmanagedType.Interface)> ByRef ppshv As IShellView) As Integer
        <PreserveSig>
        Function OnViewWindowActive(<MarshalAs(UnmanagedType.Interface)> pshv As IShellView) As Integer
        <PreserveSig>
        Function SetToolbarItems(lpButtons As IntPtr, nButtons As UInt32, uFlags As UInt32) As Integer
    End Interface
End Namespace

