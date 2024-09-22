Imports System.Runtime.InteropServices

<StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)>
Public Structure DSOBJECT
    Public dwFlags As Integer
    Public dwProviderFlags As Integer
    Public offsetName As Integer
    Public offsetClass As Integer
End Structure