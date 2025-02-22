Imports System.Runtime.InteropServices

Namespace Interop.DragDrop
    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)>
    Public Structure DROPDESCRIPTION
        Public type As DROPIMAGETYPE
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=260)> Public szMessage As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=260)> Public szInsert As String
    End Structure
End Namespace

