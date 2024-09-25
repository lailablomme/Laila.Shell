Imports System.Runtime.InteropServices

<StructLayout(LayoutKind.Sequential)>
Public Structure DEFCONTEXTMENU
    Public hwnd As IntPtr
    Public pcmcb As IntPtr
    Public pidlFolder As IntPtr
    Public psf As IntPtr
    Public cidl As UInteger
    Public apidl As IntPtr
    Public punkAssociationInfo As IntPtr
    Public cKeys As UInteger
    Public aKeys As IntPtr
End Structure
