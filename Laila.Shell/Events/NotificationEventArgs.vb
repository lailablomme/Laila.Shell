Imports Laila.Shell.Interop.Items

Namespace Events
    Public Class NotificationEventArgs
        Inherits EventArgs

        Public Property Item1 As Item
        Public Property Item2 As Item
        Public Property [Event] As SHCNE
        Public Property IsHandled1 As Boolean
        Public Property IsHandled2 As Boolean
    End Class
End Namespace