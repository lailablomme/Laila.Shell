Imports System.Runtime.InteropServices

Namespace Interop.ContextMenu
    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Auto)>
    Public Structure MENUITEMINFO
        Public cbSize As UInteger
        Public fMask As UInteger
        Public fType As UInteger
        Public fState As MFS
        Public wID As Integer
        Public hSubMenu As IntPtr
        Public hbmpChecked As IntPtr
        Public hbmpUnchecked As IntPtr
        Public dwItemData As IntPtr
        Public dwTypeData As IntPtr
        Public cch As UInteger
        Public hbmpItem As IntPtr
    End Structure
End Namespace

