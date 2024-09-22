Imports System.Runtime.InteropServices

<StructLayout(LayoutKind.Sequential)>
Public Structure MENUITEMINFO
    Public cbSize As UInteger
    Public fMask As UInteger
    Public fType As UInteger
    Public fState As UInteger
    Public wID As UInt32
    Public hSubMenu As IntPtr
    Public hbmpChecked As IntPtr
    Public hbmpUnchecked As IntPtr
    Public dwItemData As IntPtr
    Public dwTypeData As String
    Public cch As UInteger
    Public hbmpItem As IntPtr
End Structure
