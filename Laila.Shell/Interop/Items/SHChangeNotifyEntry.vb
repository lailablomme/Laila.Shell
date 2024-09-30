Imports System.Runtime.InteropServices

<StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Auto)>
Public Structure SHChangeNotifyEntry
    Public pIdl As IntPtr
    Public Recursively As Boolean
End Structure