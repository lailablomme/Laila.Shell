Imports System.Runtime.InteropServices
Imports System.Windows

<StructLayout(LayoutKind.Sequential)>
Public Structure DRAWITEMSTRUCT
    Public CtlType As UInteger
    Public CtlID As UInteger
    Public itemID As UInteger
    Public itemAction As UInteger
    Public itemState As UInteger
    Public hwndItem As IntPtr
    Public hDC As IntPtr
    Public rcItem As RECT
    Public itemData As IntPtr
End Structure