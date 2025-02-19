Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Controls.Primitives
Imports System.Windows.Input
Imports Laila.Shell.Helpers

Namespace Controls
    Public Class ContextMenu
        Inherits System.Windows.Controls.ContextMenu

        Public Shared ReadOnly ShowButtonsTopOrBottomProperty As DependencyProperty = DependencyProperty.Register("ShowButtonsTopOrBottom", GetType(TopOrBottom),
            GetType(ContextMenu), New FrameworkPropertyMetadata(TopOrBottom.Unset, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))

        Private _buttonsTop As StackPanel
        Private _buttonsBottom As StackPanel
        Private _scrollViewer As ScrollViewer
        Private _list As List(Of ButtonBase) = New List(Of ButtonBase)()
        Private _isMade As Boolean

        Shared Sub New()
            DefaultStyleKeyProperty.OverrideMetadata(GetType(ContextMenu), New FrameworkPropertyMetadata(GetType(ContextMenu)))
        End Sub

        Public Overrides Sub OnApplyTemplate()
            MyBase.OnApplyTemplate()

            _buttonsTop = Template.FindName("PART_ButtonsTop", Me)
            _buttonsBottom = Template.FindName("PART_ButtonsBottom", Me)
            _scrollViewer = Template.FindName("PART_ScrollViewer", Me)
        End Sub

        Public Overridable Async Function Make() As Task
            _isMade = True
        End Function

        Protected Overrides Async Sub OnOpened(e As RoutedEventArgs)
            If Not _isMade Then
                Await Me.Make()
            End If

            MyBase.OnOpened(e)

            UIHelper.OnUIThread(
                Sub()
                    If _list.Count > 0 Then
                        Dim buttonStyle As Style = FindResource("lailaShell_ContextMenuButtonStyle")
                        Dim toggleButtonStyle As Style = FindResource("lailaShell_ContextMenuToggleButtonStyle")
                        If Not PresentationSource.FromVisual(_buttonsTop) Is Nothing _
                            AndAlso Not PresentationSource.FromVisual(_buttonsBottom) Is Nothing Then
                            Dim ptButtonsTop As Point = _buttonsTop.PointToScreen(New Point(0, 0))
                            Dim ptButtonsBottom As Point = _buttonsBottom.PointToScreen(New Point(0, 0))
                            Dim ptMouse As Point = Application.Current.MainWindow.PointToScreen(Mouse.GetPosition(Application.Current.MainWindow))

                            If Math.Abs(ptMouse.Y - ptButtonsTop.Y) < Math.Abs(ptMouse.Y - ptButtonsBottom.Y) Then
                                For Each button In _list
                                    If TypeOf button Is Button AndAlso Not buttonStyle Is Nothing Then button.Style = buttonStyle
                                    If TypeOf button Is ToggleButton AndAlso Not toggleButtonStyle Is Nothing Then button.Style = toggleButtonStyle
                                    If Not button.Parent Is Nothing Then
                                        CType(button.Parent, Panel).Children.Remove(button)
                                    End If
                                    _buttonsTop.Children.Add(button)
                                Next
                                Me.ShowButtonsTopOrBottom = TopOrBottom.Top
                            Else
                                For Each button In _list
                                    If TypeOf button Is Button AndAlso Not buttonStyle Is Nothing Then button.Style = buttonStyle
                                    If TypeOf button Is ToggleButton AndAlso Not toggleButtonStyle Is Nothing Then button.Style = toggleButtonStyle
                                    If Not button.Parent Is Nothing Then
                                        CType(button.Parent, Panel).Children.Remove(button)
                                    End If
                                    _buttonsBottom.Children.Add(button)
                                Next
                                Me.ShowButtonsTopOrBottom = TopOrBottom.Bottom
                            End If
                        Else
                            Me.ShowButtonsTopOrBottom = TopOrBottom.None
                        End If
                    Else
                        Me.ShowButtonsTopOrBottom = TopOrBottom.None
                    End If
                End Sub, Threading.DispatcherPriority.ContextIdle)
        End Sub

        Public Property ShowButtonsTopOrBottom As TopOrBottom
            Get
                Return GetValue(ShowButtonsTopOrBottomProperty)
            End Get
            Set(ByVal value As TopOrBottom)
                SetCurrentValue(ShowButtonsTopOrBottomProperty, value)
            End Set
        End Property

        Public ReadOnly Property Buttons As List(Of ButtonBase)
            Get
                Return _list
            End Get
        End Property

        Public Enum TopOrBottom
            Unset
            Both
            Top
            Bottom
            None
        End Enum
    End Class
End Namespace