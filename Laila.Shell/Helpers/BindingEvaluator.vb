Imports System.Windows
Imports System.Windows.Data

Namespace Helpers
    Public Class BindingEvaluator
        Inherits FrameworkElement

        Public Shared ReadOnly ValueProperty As DependencyProperty = DependencyProperty.Register("Value", GetType(Object), GetType(BindingEvaluator), New FrameworkPropertyMetadata(Nothing))

        Private _valueBinding As Binding

        Public Sub New(ByVal binding As Binding)
            _valueBinding = binding
        End Sub

        Public Function Evaluate(ByVal dataItem As Object) As Object
            Me.DataContext = dataItem
            SetBinding(ValueProperty, _valueBinding)
            Return GetValue(ValueProperty)
        End Function
    End Class
End Namespace