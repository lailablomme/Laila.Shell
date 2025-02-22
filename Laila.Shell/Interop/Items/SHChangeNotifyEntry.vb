Imports System.Runtime.InteropServices

Namespace Interop.Items
    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Auto)>
    Public Structure SHChangeNotifyEntry
        Public pIdl As IntPtr
        Public Recursively As Boolean
    End Structure
End Namespace
