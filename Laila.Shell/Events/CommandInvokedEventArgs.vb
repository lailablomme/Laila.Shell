Namespace Events
    Public Class CommandInvokedEventArgs
        Inherits EventArgs

        Public Property Id As Integer
        Public Property Verb As String
        Public Property IsChecked As Boolean
        Public Property IsHandled As Boolean
    End Class
End Namespace