Namespace Interfaces
    Public Interface INotify
        Sub SubscribeToNotifications(item As IProcessNotifications)
        Sub UnsubscribeFromNotifications(item As IProcessNotifications)
    End Interface
End Namespace