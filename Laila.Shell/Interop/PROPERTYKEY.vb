Imports System.Runtime.InteropServices

<StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)>
Public Structure PROPERTYKEY
    Public fmtid As Guid
    Public pid As Integer

    Friend Overloads Function Equals(propertyKey As PROPERTYKEY) As Boolean
        Return propertyKey.fmtid.Equals(fmtid) AndAlso propertyKey.pid.Equals(pid)
    End Function
End Structure
