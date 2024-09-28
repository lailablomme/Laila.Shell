Imports System.Runtime.InteropServices

<StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode, Size:=1044)>
Public Structure DROPDESCRIPTION
    Public type As DropImageType
    <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=260)>
    Public szMessage As String
    <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=260)>
    Public szInsert As String
End Structure