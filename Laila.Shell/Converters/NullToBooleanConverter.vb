Imports System.Globalization
Imports System.Windows.Data

Namespace Converters
    Public Class NullToBooleanConverter
        Implements IValueConverter

        Public Property Invert As Boolean = False

        Public Function Convert(value As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements IValueConverter.Convert
            Dim isNull = value Is Nothing
            Return If(Invert, Not isNull, isNull)
        End Function

        Public Function ConvertBack(value As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements IValueConverter.ConvertBack
            Throw New NotImplementedException()
        End Function
    End Class
End Namespace