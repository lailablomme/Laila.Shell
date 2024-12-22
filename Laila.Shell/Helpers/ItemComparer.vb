Imports System.ComponentModel

Namespace Helpers
    Public Class ItemComparer
        Implements IComparer

        Private _groupByPropertyName As String
        Private _sortPropertyName As String
        Private _direction As ListSortDirection
        Private _stringComparer As StringComparer = New StringComparer()

        Public Sub New(groupByPropertyName As String, sortPropertyName As String, direction As ListSortDirection)
            _groupByPropertyName = groupByPropertyName
            _sortPropertyName = sortPropertyName
            _direction = direction
        End Sub

        Public Function Compare(x As Object, y As Object) As Integer Implements IComparer.Compare
            If Not String.IsNullOrWhiteSpace(_groupByPropertyName) Then
                Dim prop1ValueX As Object = getPropertyText(x, _groupByPropertyName)
                Dim prop1ValueY As Object = getPropertyText(y, _groupByPropertyName)

                Dim result As Integer
                If TypeOf prop1ValueX Is String Then
                    result = _stringComparer.Compare(prop1ValueX, prop1ValueY)
                Else
                    result = Comparer.Default.Compare(prop1ValueX, prop1ValueY)
                End If
                If result <> 0 Then
                    ' Adjust for sort direction
                    result *= If(_direction = ListSortDirection.Ascending, 1, -1)
                    Return result
                End If
            End If

            If Not String.IsNullOrWhiteSpace(_sortPropertyName) Then
                Dim prop2ValueX As Object = getPropertyValue(x, _sortPropertyName)
                Dim prop2ValueY As Object = getPropertyValue(y, _sortPropertyName)

                Dim result As Integer
                If TypeOf prop2ValueX Is String Then
                    result = _stringComparer.Compare(prop2ValueX, prop2ValueY)
                Else
                    result = Comparer.Default.Compare(prop2ValueX, prop2ValueY)
                End If
                ' Adjust for sort direction
                result *= If(_direction = ListSortDirection.Ascending, 1, -1)
                Return result
            End If

            Return 1
        End Function

        Private Function getPropertyValue(item As Item, propertyName As String) As Object
            If propertyName = "ItemNameDisplaySortValue" Then
                Return item.ItemNameDisplaySortValue
            ElseIf propertyName.Substring(0, 18) = "PropertiesByKeyAsT" Then
                Return item.PropertiesByKeyAsText(propertyName.Substring(22, propertyName.LastIndexOf("]") - 22))?.Value
            ElseIf propertyName.Substring(0, 18) = "PropertiesByCanoni" Then
                Return item.PropertiesByCanonicalName(propertyName.Substring(26, propertyName.LastIndexOf("]") - 26))?.Value
            Else
                Throw New NotSupportedException()
            End If
        End Function

        Private Function getPropertyText(item As Item, propertyName As String) As Object
            If propertyName = "ItemNameDisplaySortValue" Then
                Return item.ItemNameDisplaySortValue
            ElseIf propertyName.Substring(0, 18) = "PropertiesByKeyAsT" Then
                Return item.PropertiesByKeyAsText(propertyName.Substring(22, propertyName.LastIndexOf("]") - 22))?.Text
            ElseIf propertyName.Substring(0, 18) = "PropertiesByCanoni" Then
                Return item.PropertiesByCanonicalName(propertyName.Substring(26, propertyName.LastIndexOf("]") - 26))?.Text
            Else
                Throw New NotSupportedException()
            End If
        End Function
    End Class
End Namespace