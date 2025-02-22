Imports System.Runtime.InteropServices

Namespace Interop.SevenZip
    Public Class Functions
        <DllImport("7z.dll", CallingConvention:=CallingConvention.StdCall)>
        Public Shared Function CreateObject(<[In]> ByRef clsID As Guid, <[In]> ByRef iid As Guid, <Out, MarshalAs(UnmanagedType.Interface)> ByRef outObject As Object) As Integer
        End Function
    End Class
End Namespace