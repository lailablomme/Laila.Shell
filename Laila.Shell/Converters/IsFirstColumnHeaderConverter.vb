Imports System.Globalization
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Data

Namespace Converters
    Public Class IsFirstColumnHeaderConverter
        Implements IMultiValueConverter

        Public Function Convert(values() As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object _
            Implements IMultiValueConverter.Convert

            Dim item As GridViewColumnHeader = TryCast(values(0), GridViewColumnHeader)
            Dim listView As ListView = TryCast(values(1), ListView)

            Return item.Column?.Equals(CType(listView.View, GridView).Columns.FirstOrDefault())
        End Function

        Public Function ConvertBack(value As Object, targetTypes() As Type, parameter As Object, culture As CultureInfo) As Object() _
            Implements IMultiValueConverter.ConvertBack
            Throw New NotImplementedException()
        End Function
    End Class
End Namespace