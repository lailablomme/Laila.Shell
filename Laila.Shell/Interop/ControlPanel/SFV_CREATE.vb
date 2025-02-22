Imports System.Runtime.InteropServices
Imports Laila.Shell.Interop.Folders

Namespace Interop.ControlPanel
    <StructLayout(LayoutKind.Sequential)>
    Public Structure SFV_CREATE
        Public cbSize As UInt32
        Public pshf As IShellFolder    ' Pointer to IShellFolder
        Public psvOuter As IShellView  ' Pointer to IShellView (could be null)
        Public psfvcb As IntPtr ' Pointer to IShellFolderViewCB (callback interface, could be null)
    End Structure
End Namespace
