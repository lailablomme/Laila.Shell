Imports System.ComponentModel

Namespace Helpers
    Public Class TreeViewComparer
        Implements IComparer

        Private ReadOnly _propertyName As String
        Private ReadOnly _comparer As StringComparer

        Public Sub New(propertyName As String)
            _propertyName = propertyName
            _comparer = New StringComparer()
        End Sub

        Public Function Compare(x As Object, y As Object) As Integer Implements IComparer.Compare
            Dim propertyInfo = x.GetType().GetProperty(_propertyName)
            If propertyInfo IsNot Nothing Then
                Dim valueX = CType(propertyInfo.GetValue(x, Nothing), String)
                Dim valueY = CType(propertyInfo.GetValue(y, Nothing), String)
                Return _comparer.Compare(valueX, valueY)
            End If
            Return 0
        End Function
    End Class
End Namespace