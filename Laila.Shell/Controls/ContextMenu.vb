Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Controls.Primitives
Imports System.Windows.Input

Namespace Controls
    Public Class ContextMenu
        Inherits Windows.Controls.ContextMenu

        Public Shared ReadOnly ShowButtonsTopOrBottomProperty As DependencyProperty = DependencyProperty.Register("ShowButtonsTopOrBottom", GetType(Boolean?),
            GetType(ContextMenu), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))

        Private _buttonsTop As StackPanel
        Private _buttonsBottom As StackPanel
        Private _list As List(Of Button) = New List(Of Button)()

        Shared Sub New()
            DefaultStyleKeyProperty.OverrideMetadata(GetType(ContextMenu), New FrameworkPropertyMetadata(GetType(ContextMenu)))
        End Sub

        Public Overrides Sub OnApplyTemplate()
            MyBase.OnApplyTemplate()

            _buttonsTop = Template.FindName("PART_ButtonsTop", Me)
            _buttonsBottom = Template.FindName("PART_ButtonsBottom", Me)

        End Sub

        Protected Overrides Sub OnOpened(e As RoutedEventArgs)
            MyBase.OnOpened(e)

            Dim ptButtonsTop As Point = _buttonsTop.PointToScreen(New Point(0, 0))
            Dim ptButtonsBottom As Point = _buttonsBottom.PointToScreen(New Point(0, 0))
            Dim ptMouse As Point = Application.Current.MainWindow.PointToScreen(Mouse.GetPosition(Application.Current.MainWindow))

            If Math.Abs(ptMouse.Y - ptButtonsTop.Y) < Math.Abs(ptMouse.Y - ptButtonsBottom.Y) Then
                For Each button In _list
                    _buttonsTop.Children.Add(button)
                Next
                Me.ShowButtonsTopOrBottom = True
            Else
                For Each button In _list
                    _buttonsBottom.Children.Add(button)
                Next
                Me.ShowButtonsTopOrBottom = False
            End If
        End Sub

        Public Property ShowButtonsTopOrBottom As Boolean?
            Get
                Return GetValue(ShowButtonsTopOrBottomProperty)
            End Get
            Set(ByVal value As Boolean?)
                SetCurrentValue(ShowButtonsTopOrBottomProperty, value)
            End Set
        End Property

        Public ReadOnly Property Buttons As List(Of Button)
            Get
                Return _list
            End Get
        End Property
    End Class
End Namespace