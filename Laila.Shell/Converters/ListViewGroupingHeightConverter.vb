Imports System.Globalization
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Data

Namespace Converters
    Public Class ListViewGroupingHeightConverter
        Implements IMultiValueConverter

        Public Function Convert(values() As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements IMultiValueConverter.Convert
            Dim stackPanel As StackPanel = values(0)
            stackPanel.Measure(New Windows.Size(Double.PositiveInfinity, Double.PositiveInfinity))
            Return System.Convert.ToDouble(values(1) - stackPanel.DesiredSize.Height - 25)
        End Function

        Public Function ConvertBack(value As Object, targetTypes() As Type, parameter As Object, culture As CultureInfo) As Object() Implements IMultiValueConverter.ConvertBack
            Throw New NotImplementedException()
        End Function
    End Class
End Namespace