Imports System.Runtime.InteropServices
Imports System.Windows

<StructLayout(LayoutKind.Sequential)>
Public Structure IMAGEINFO
    Public hbmImage As IntPtr
    Public hbmMask As IntPtr
    Public Unused1 As Integer
    Public Unused2 As Integer
    Public rcImage As RECT
End Structure
