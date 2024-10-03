Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Controls.Primitives

Namespace Controls
    Public Class ContextMenu
        Inherits Windows.Controls.ContextMenu

        Private _buttonsTop As StackPanel
        Private _buttonsBottom As StackPanel
        Private _listTop As List(Of Button) = New List(Of Button)()
        Private _listBottom As List(Of Button) = New List(Of Button)()

        Shared Sub New()
            DefaultStyleKeyProperty.OverrideMetadata(GetType(ContextMenu), New FrameworkPropertyMetadata(GetType(ContextMenu)))
        End Sub

        Public Overrides Sub OnApplyTemplate()
            MyBase.OnApplyTemplate()

            _buttonsTop = Template.FindName("PART_ButtonsTop", Me)
            _buttonsBottom = Template.FindName("PART_ButtonsBottom", Me)

            For Each button In _listTop
                _buttonsTop.Children.Add(button)
            Next
            For Each button In _listBottom
                _buttonsBottom.Children.Add(button)
            Next
        End Sub

        Public ReadOnly Property ButtonsTop As List(Of Button)
            Get
                Return _listTop
            End Get
        End Property

        Public ReadOnly Property ButtonsBottom As List(Of Button)
            Get
                Return _listBottom
            End Get
        End Property
    End Class
End Namespace