Imports Laila.Shell.Events

Namespace Interfaces
    Public Interface IProcessNotifications
        Property IsProcessingNotifications As Boolean
        Property NotificationThreadId As Integer?
        Sub ProcessNotification(ByVal e As NotificationEventArgs)
    End Interface
End Namespace