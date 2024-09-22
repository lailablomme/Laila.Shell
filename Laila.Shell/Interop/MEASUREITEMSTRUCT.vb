Imports System.Runtime.InteropServices

<StructLayout(LayoutKind.Sequential)>
Public Structure MEASUREITEMSTRUCT
    Public CtlType As Integer
    Public CtlID As Integer
    Public itemID As Integer
    Public itemWidth As Integer
    Public itemHeight As Integer
    Public itemData As IntPtr
End Structure