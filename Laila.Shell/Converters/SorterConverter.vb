Imports System.ComponentModel
Imports System.Windows.Data

Namespace Converters
    Public Class SorterConverter
        Implements IValueConverter

        Public Function Convert(value As Object, targetType As Type, parameter As Object, culture As Globalization.CultureInfo) As Object Implements IValueConverter.Convert
            Dim list As IList = value
            If list Is Nothing Then
                Return Nothing
            Else
                Dim view As ListCollectionView = New ListCollectionView(list)
                view.SortDescriptions.Add(New SortDescription(parameter.ToString(), ListSortDirection.Ascending))
                view.Filter = Function(obj As Object) As Boolean
                                  Return TypeOf obj Is Folder
                              End Function
                Return view
            End If
        End Function

        Public Function ConvertBack(value As Object, targetType As Type, parameter As Object, culture As Globalization.CultureInfo) As Object Implements IValueConverter.ConvertBack
            Throw New NotImplementedException()
        End Function
    End Class
End Namespace