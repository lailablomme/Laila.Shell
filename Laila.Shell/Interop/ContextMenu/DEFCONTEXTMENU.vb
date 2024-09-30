Imports System.Runtime.InteropServices

<StructLayout(LayoutKind.Sequential)>
Public Structure DEFCONTEXTMENU
    Public hwnd As IntPtr
    <MarshalAs(UnmanagedType.Interface)> Public pcmcb As IContextMenuCB
    Public pidlFolder As IntPtr
    <MarshalAs(UnmanagedType.Interface)> Public psf As IShellFolder
    Public cidl As UInteger
    Public apidl As IntPtr
    Public punkAssociationInfo As IntPtr
    Public cKeys As UInteger
    Public aKeys As IntPtr
End Structure
