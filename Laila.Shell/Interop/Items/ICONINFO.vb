Imports System.Runtime.InteropServices

Namespace Interop.Items
    <StructLayout(LayoutKind.Sequential)>
    Public Structure ICONINFO
        Public fIcon As Boolean
        Public xHotspot As Int32
        Public yHotspot As Int32
        Public hbmMask As IntPtr
        Public hbmColor As IntPtr
    End Structure
End Namespace

