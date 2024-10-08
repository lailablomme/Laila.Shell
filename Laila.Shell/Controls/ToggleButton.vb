Imports System.Windows
Imports System.Windows.Input
Imports Laila.Shell.Helpers

Namespace Controls
    Public Class ToggleButton
        Inherits System.Windows.Controls.Control

        Public Shared ReadOnly IsCheckedProperty As DependencyProperty = DependencyProperty.Register("IsChecked", GetType(Boolean), GetType(ToggleButton), New FrameworkPropertyMetadata(False, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnIsCheckedChanged))

        Public Event Checked(sender As Object, e As EventArgs)
        Public Event Unchecked(sender As Object, e As EventArgs)

        Shared Sub New()
            DefaultStyleKeyProperty.OverrideMetadata(GetType(ToggleButton), New FrameworkPropertyMetadata(GetType(ToggleButton)))
        End Sub

        Public Sub New()
            AddHandler Me.MouseUp,
                Sub(s As Object, e As MouseButtonEventArgs)
                    UIHelper.OnUIThreadAsync(
                        Sub()
                            Me.IsChecked = Not Me.IsChecked
                            If Me.IsChecked Then
                                RaiseEvent Checked(Me, New EventArgs())
                            Else
                                RaiseEvent Unchecked(Me, New EventArgs())
                            End If
                        End Sub, Threading.DispatcherPriority.ContextIdle)
                    e.Handled = True
                End Sub
        End Sub

        Public Property IsChecked As Boolean
            Get
                Return GetValue(IsCheckedProperty)
            End Get
            Set(ByVal value As Boolean)
                SetCurrentValue(IsCheckedProperty, value)
            End Set
        End Property

        Shared Sub OnIsCheckedChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            Dim i = 9
        End Sub
    End Class
End Namespace