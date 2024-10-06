Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Controls.Primitives
Imports System.Windows.Input
Imports Laila.Shell.Helpers

Namespace Controls
    Public Class ContextMenu
        Inherits Windows.Controls.ContextMenu

        Public Shared ReadOnly ShowButtonsTopOrBottomProperty As DependencyProperty = DependencyProperty.Register("ShowButtonsTopOrBottom", GetType(Boolean?),
            GetType(ContextMenu), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))

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

            Dim style As Style = FindResource("lailaShell_ContextMenuButtonStyle")
            Dim ptButtonsTop As Point = _buttonsTop.PointToScreen(New Point(0, 0))
            Dim ptButtonsBottom As Point = _buttonsBottom.PointToScreen(New Point(0, 0))
            Dim ptMouse As Point = Application.Current.MainWindow.PointToScreen(Mouse.GetPosition(Application.Current.MainWindow))

            If Math.Abs(ptMouse.Y - ptButtonsTop.Y) < Math.Abs(ptMouse.Y - ptButtonsBottom.Y) Then
                For Each button In _list
                    If Not style Is Nothing Then button.Style = style
                    _buttonsTop.Children.Add(button)
                Next
                Me.ShowButtonsTopOrBottom = True
            Else
                For Each button In _list
                    If Not style Is Nothing Then button.Style = style
                    _buttonsBottom.Children.Add(button)
                Next
                Me.ShowButtonsTopOrBottom = False
            End If

            _buttonsTop.Measure(New Size(1000, 1000))
            _buttonsBottom.Measure(New Size(1000, 1000))

            _scrollViewer.MaxHeight = UIHelper.GetScreenSize(_scrollViewer, True).Height _
                - _buttonsTop.DesiredSize.Height - _buttonsBottom.DesiredSize.Height - 16
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