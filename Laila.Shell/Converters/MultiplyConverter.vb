Imports System.Globalization
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Data

Namespace Converters
    Public Class MultiplyConverter
        Implements IMultiValueConverter

        Public Function Convert(values() As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements IMultiValueConverter.Convert
            Dim b As Border = values(0)
            b.Measure(New Size(Double.PositiveInfinity, Double.PositiveInfinity))
            Return New GridLength(System.Convert.ToDouble(b.DesiredSize.Height) * System.Convert.ToDouble(values(1)), GridUnitType.Pixel)
        End Function

        Public Function ConvertBack(value As Object, targetTypes() As Type, parameter As Object, culture As CultureInfo) As Object() Implements IMultiValueConverter.ConvertBack
            Throw New NotImplementedException()
        End Function
    End Class
End Namespace