Imports System.Globalization
Imports System.Windows
Imports System.Windows.Data
Imports System.Windows.Media

Namespace Converters
    Public Class OpacityConverter
        Implements IValueConverter

        ' Parameter: the desired opacity (0.0–1.0)
        Public Function Convert(value As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements IValueConverter.Convert
            Dim opacity As Double = 1.0

            If parameter IsNot Nothing Then
                Double.TryParse(parameter.ToString(), NumberStyles.Any, CultureInfo.InvariantCulture, opacity)
                opacity = Math.Max(0.0, Math.Min(1.0, opacity))
            End If
            If TypeOf value Is Color Then
                Dim color = CType(value, Color)
                Return Color.FromArgb(CByte(opacity * 255), color.R, color.G, color.B)
            ElseIf TypeOf value Is SolidColorBrush Then
                Dim brush = CType(value, SolidColorBrush)
                Dim color = brush.Color
                Return New SolidColorBrush(Color.FromArgb(CByte(opacity * 255), color.R, color.G, color.B))
            Else
                Return DependencyProperty.UnsetValue
            End If
        End Function

        Public Function ConvertBack(value As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements IValueConverter.ConvertBack
            Throw New NotSupportedException()
        End Function
    End Class
End Namespace