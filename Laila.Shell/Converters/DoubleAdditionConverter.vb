﻿Imports System.Globalization
Imports System.Windows.Data

Namespace Converters
    Public Class DoubleAdditionConverter
        Implements IValueConverter

        Public Function Convert(value As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements IValueConverter.Convert
            Return System.Convert.ToDouble(value, CultureInfo.InvariantCulture) + System.Convert.ToDouble(parameter, CultureInfo.InvariantCulture)
        End Function

        Public Function ConvertBack(value As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements IValueConverter.ConvertBack
            Throw New NotImplementedException()
        End Function
    End Class
End Namespace