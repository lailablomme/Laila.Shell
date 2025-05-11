Imports System.Globalization
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Data

Namespace Converters
    Public Class MultiplyConverterGridLength
        Implements IMultiValueConverter

        Public Function Convert(values() As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements IMultiValueConverter.Convert
            Return New GridLength(System.Convert.ToDouble(values(0)) * System.Convert.ToDouble(values(1), CultureInfo.InvariantCulture), GridUnitType.Pixel)
        End Function

        Public Function ConvertBack(value As Object, targetTypes() As Type, parameter As Object, culture As CultureInfo) As Object() Implements IMultiValueConverter.ConvertBack
            Throw New NotImplementedException()
        End Function
    End Class

    Public Class MultiplyConverter
        Implements IMultiValueConverter

        Public Function Convert(values() As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements IMultiValueConverter.Convert
            Return System.Convert.ToDouble(values(0)) * System.Convert.ToDouble(values(1))
        End Function

        Public Function ConvertBack(value As Object, targetTypes() As Type, parameter As Object, culture As CultureInfo) As Object() Implements IMultiValueConverter.ConvertBack
            Throw New NotImplementedException()
        End Function
    End Class

    Public Class SubstractionConverter
        Implements IMultiValueConverter

        Public Function Convert(values() As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements IMultiValueConverter.Convert
            Return System.Convert.ToDouble(values(0)) - System.Convert.ToDouble(values(1))
        End Function

        Public Function ConvertBack(value As Object, targetTypes() As Type, parameter As Object, culture As CultureInfo) As Object() Implements IMultiValueConverter.ConvertBack
            Throw New NotImplementedException()
        End Function
    End Class
End Namespace