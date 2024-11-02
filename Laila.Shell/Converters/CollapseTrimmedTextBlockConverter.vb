Imports System.Globalization
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Data

Namespace Converters
    Public Class CollapseTrimmedTextBlockConverter
        Implements IMultiValueConverter

        Public Function Convert(values() As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements IMultiValueConverter.Convert
            Dim textBlock As TextBlock = values(0)
            textBlock.Measure(New Windows.Size(Double.PositiveInfinity, Double.PositiveInfinity))
            Return If(textBlock.ActualWidth <= textBlock.DesiredSize.Width, Visibility.Hidden, Visibility.Visible)
        End Function

        Public Function ConvertBack(value As Object, targetTypes() As Type, parameter As Object, culture As CultureInfo) As Object() Implements IMultiValueConverter.ConvertBack
            Throw New NotImplementedException()
        End Function
    End Class
End Namespace