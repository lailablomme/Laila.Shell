Imports System.Runtime.InteropServices

<StructLayout(LayoutKind.Sequential)>
Public Structure SFV_CREATE
    Public cbSize As UInt32
    Public pshf As IShellFolder    ' Pointer to IShellFolder
    Public psvOuter As IShellView  ' Pointer to IShellView (could be null)
    Public psfvcb As IntPtr ' Pointer to IShellFolderViewCB (callback interface, could be null)
End Structure