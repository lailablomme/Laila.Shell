Namespace Helpers
    Public Class TreeViewStringComparer
        Inherits StringComparer

        Public Overrides Function Compare(x As String, y As String) As Integer
            Return String.Compare(x, y, StringComparison.OrdinalIgnoreCase)
        End Function

        Public Overrides Function Equals(x As String, y As String) As Boolean
            Return String.Equals(x, y, StringComparison.OrdinalIgnoreCase)
        End Function

        Public Overrides Function GetHashCode(obj As String) As Integer
            Return obj.GetHashCode()
        End Function
    End Class
End Namespace