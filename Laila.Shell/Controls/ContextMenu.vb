Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Controls.Primitives
Imports System.Windows.Input
Imports Laila.Shell.Helpers

Namespace Controls
    Public Class ContextMenu
        Inherits Windows.Controls.ContextMenu

        Public Shared ReadOnly ShowButtonsTopOrBottomProperty As DependencyProperty = DependencyProperty.Register("ShowButtonsTopOrBottom", GetType(TopOrBottom),
            GetType(ContextMenu), New FrameworkPropertyMetadata(TopOrBottom.Both, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))

        Private _buttonsTop As StackPanel
        Private _buttonsBottom As StackPanel
        Private _scrollViewer As ScrollViewer
        Private _list As List(Of Button) = New List(Of Button)()

        Shared Sub New()
            DefaultStyleKeyProperty.OverrideMetadata(GetType(ContextMenu), New FrameworkPropertyMetadata(GetType(ContextMenu)))
        End Sub

        Public Overrides Sub OnApplyTemplate()
            MyBase.OnApplyTemplate()

            _buttonsTop = Template.FindName("PART_ButtonsTop", Me)
            _buttonsBottom = Template.FindName("PART_ButtonsBottom", Me)
            _scrollViewer = Template.FindName("PART_ScrollViewer", Me)
        End Sub

        Protected Overrides Sub OnOpened(e As RoutedEventArgs)
            MyBase.OnOpened(e)

            If _list.Count > 0 Then
                Dim style As Style = FindResource("lailaShell_ContextMenuButtonStyle")
                Dim ptButtonsTop As Point = _buttonsTop.PointToScreen(New Point(0, 0))
                Dim ptButtonsBottom As Point = _buttonsBottom.PointToScreen(New Point(0, 0))
                Dim ptMouse As Point = Application.Current.MainWindow.PointToScreen(Mouse.GetPosition(Application.Current.MainWindow))

                If Math.Abs(ptMouse.Y - ptButtonsTop.Y) < Math.Abs(ptMouse.Y - ptButtonsBottom.Y) Then
                    For Each button In _list
                        If Not style Is Nothing Then button.Style = style
                        _buttonsTop.Children.Add(button)
                    Next
                    Me.ShowButtonsTopOrBottom = TopOrBottom.Top
                Else
                    For Each button In _list
                        If Not style Is Nothing Then button.Style = style
                        _buttonsBottom.Children.Add(button)
                    Next
                    Me.ShowButtonsTopOrBottom = TopOrBottom.Bottom
                End If
            Else
                Me.ShowButtonsTopOrBottom = TopOrBottom.None
            End If
        End Sub

        Public Property ShowButtonsTopOrBottom As TopOrBottom
            Get
                Return GetValue(ShowButtonsTopOrBottomProperty)
            End Get
            Set(ByVal value As TopOrBottom)
                SetCurrentValue(ShowButtonsTopOrBottomProperty, value)
            End Set
        End Property

        Public ReadOnly Property Buttons As List(Of Button)
            Get
                Return _list
            End Get
        End Property

        Public Enum TopOrBottom
            Both
            Top
            Bottom
            None
        End Enum
    End Class
End Namespace