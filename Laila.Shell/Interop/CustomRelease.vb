Imports System.Runtime.InteropServices

<ComVisible(True)>
Public Class CustomRelease
    Implements IUnknown

    Private refCount As Integer = 0

    Public Sub New()
        refCount = 1
    End Sub

    Public Function AddRef() As Integer Implements IUnknown.AddRef
        refCount += 1
        Return refCount
    End Function

    Public Function Release() As Integer Implements IUnknown.Release
        refCount -= 1
        If refCount = 0 Then
            ' Custom cleanup code here
            Marshal.ReleaseComObject(Me)
        End If
        Return refCount
    End Function

    Public Function QueryInterface(ByRef riid As Guid, ByRef ppvObject As IntPtr) As Integer Implements IUnknown.QueryInterface
        ' Implement QueryInterface logic here
        Return 0
    End Function
End Class