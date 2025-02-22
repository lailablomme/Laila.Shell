Imports Laila.Shell.Interop
Imports Microsoft.Win32.SafeHandles

Namespace Helpers
    Public Class SafeCursorHandle
        Inherits SafeHandleZeroOrMinusOneIsInvalid

        Public Sub New(ptr As IntPtr)
            MyBase.New(True)

            SetHandle(ptr)
        End Sub

        Protected Overrides Function ReleaseHandle() As Boolean
            If Not IntPtr.Zero.Equals(handle) Then
                Functions.DestroyCursor(handle)
                handle = IntPtr.Zero
            End If
            Return True
        End Function
    End Class
End Namespace