Imports System.Runtime.InteropServices

<StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)>
Public Structure FILEGROUPDESCRIPTORW
    Public cItems As UInteger
    <MarshalAs(UnmanagedType.ByValArray)>
    Public fd() As FILEDESCRIPTORW
End Structure