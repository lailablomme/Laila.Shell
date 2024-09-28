Imports Microsoft.Win32.SafeHandles

Public Class SafeCursorHandle
    Inherits SafeHandleZeroOrMinusOneIsInvalid

    Public Sub New(ptr As IntPtr)
        MyBase.New(True)

        SetHandle(ptr)
    End Sub

    Protected Overrides Function ReleaseHandle() As Boolean
        Return Functions.DestroyCursor(handle)
    End Function
End Class
