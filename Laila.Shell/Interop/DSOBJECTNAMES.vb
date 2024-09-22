Imports System.Runtime.InteropServices

<StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)>
Public Structure DSOBJECTNAMES
    Public clsidNamespace As Guid
    Public cItems As Integer
    <MarshalAs(UnmanagedType.ByValArray, SizeConst:=1)>
    Public aObjects() As DSOBJECT
End Structure