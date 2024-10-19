Imports System.Globalization

Namespace Converters
    Public Class RectConverter
        Implements IMultiValueConverter

        ' Convert method for combining ActualWidth and ActualHeight into a Rect
        Public Function Convert(values As Object(), targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements IMultiValueConverter.Convert
            If values IsNot Nothing AndAlso values.Length = 2 Then
                Dim width As Double = If(IsNumeric(values(0)), CDbl(values(0)), 0)
                Dim height As Double = If(IsNumeric(values(1)), CDbl(values(1)), 0)

                ' Create and return a Rect with the given width and height
                Return New Rect(0, 0, width, height)
            End If

            Return New Rect(0, 0, 0, 0)
        End Function

        ' ConvertBack method is not needed for this case
        Public Function ConvertBack(value As Object, targetTypes As Type(), parameter As Object, culture As CultureInfo) As Object() Implements IMultiValueConverter.ConvertBack
            Throw New NotImplementedException()
        End Function
    End Class
End Namespace