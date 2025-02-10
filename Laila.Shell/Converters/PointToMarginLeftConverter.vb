Imports System.Globalization
Imports System.Windows
Imports System.Windows.Data

Namespace Converters
    Public Class PointToMarginLeftConverter
        Implements IValueConverter

        Public Function Convert(value As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements IValueConverter.Convert
            If TypeOf value Is Point Then
                Dim p As Point = CType(value, Point)
                Return New Thickness(-p.X, 0, 0, 0) ' Converts Point(X, Y) -> Thickness(Left, Top, 0, 0)
            End If
            Return New Thickness(0)
        End Function

        Public Function ConvertBack(value As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements IValueConverter.ConvertBack
            If TypeOf value Is Thickness Then
                Dim t As Thickness = CType(value, Thickness)
                Return New Point(-t.Left, 0) ' Converts Thickness(Left, Top, 0, 0) -> Point(X, Y)
            End If
            Return New Point(0, 0)
        End Function
    End Class
End Namespace