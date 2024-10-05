Imports System.Runtime.InteropServices

<StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)>
Public Structure PROPERTYKEY
    Public fmtid As Guid
    Public pid As Integer

    Public Sub New(fmtid As String, pid As Integer)
        Me.fmtid = New Guid(fmtid)
        Me.pid = pid
    End Sub

    Friend Overloads Function Equals(propertyKey As PROPERTYKEY) As Boolean
        Return propertyKey.fmtid.Equals(fmtid) AndAlso propertyKey.pid.Equals(pid)
    End Function

    Public Overrides Function ToString() As String
        Return String.Format("{0}:{1}", fmtid.ToString(), pid)
    End Function
End Structure
