Imports System.Threading
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Controls.Primitives
Imports System.Windows.Input
Imports System.Windows.Media
Imports System.Windows.Shapes
Imports Laila.Shell.Helpers
Imports Microsoft.Xaml.Behaviors

Namespace Behaviors
    Public Class SelectionBehavior
        Inherits Behavior(Of ListBox)

        Public Shared ReadOnly IsSelectingProperty As DependencyProperty = DependencyProperty.Register("IsSelecting", GetType(Boolean), GetType(SelectionBehavior), New FrameworkPropertyMetadata(False, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))

        Private _listBox As ListBox
        Private _selectionRectangle As Border
        Private _canStartSelecting As Boolean
        Private _mouseDownPos As Point
        Private _mouseOriginalSourceDown As Object
        Private _control As Control
        Private _sv As ScrollViewer
        Private _headerHeight As Double
        Private _ptMin As Point
        Private _ptMax As Point
        Private _scrollTimer As Timer
        Private _panel As Panel
        Private _grid As Grid
        Private _isLoaded As Boolean

        Protected Overrides Sub OnAttached()
            MyBase.OnAttached()

            _listBox = Me.AssociatedObject
            _control = New Control() With {.Opacity = 0}

            AddHandler _listBox.Loaded,
                Sub(s As Object, e As EventArgs)
                    If Not _isLoaded Then
                        _isLoaded = True

                        _sv = UIHelper.FindVisualChildren(Of ScrollViewer)(_listBox)(0)
                        _selectionRectangle = New Border() With {
                            .BorderBrush = Brushes.SkyBlue,
                            .BorderThickness = New Thickness(1),
                            .Background = New SolidColorBrush(Color.FromArgb(75, Colors.SkyBlue.R, Colors.SkyBlue.G, Colors.SkyBlue.B)),
                            .Visibility = Visibility.Collapsed,
                            .HorizontalAlignment = HorizontalAlignment.Left,
                            .VerticalAlignment = VerticalAlignment.Top
                        }
                        _panel = _listBox.Parent
                        _grid = New Grid()
                        _grid.SetValue(Panel.ZIndexProperty, 2)
                        _grid.HorizontalAlignment = HorizontalAlignment.Left
                        _grid.VerticalAlignment = VerticalAlignment.Top
                        _grid.ClipToBounds = True
                        _grid.Children.Add(_selectionRectangle)
                        _grid.Children.Add(_control)
                        _panel.Children.Add(_grid)
                    End If
                End Sub

            AddHandler _listBox.PreviewMouseDown,
                Sub(s As Object, e As MouseButtonEventArgs)
                    If e.LeftButton = Input.MouseButtonState.Pressed AndAlso Not e.OriginalSource Is Nothing Then
                        Dim listViewItem As ListViewItem = UIHelper.GetParentOfType(Of ListViewItem)(e.OriginalSource)
                        Dim clickedItem As Item = listViewItem?.DataContext
                        If TypeOf e.OriginalSource Is TextBlock Then
                            _mouseOriginalSourceDown = e.OriginalSource
                        Else
                            _mouseOriginalSourceDown = UIHelper.GetParentOfType(Of TextBlock)(e.OriginalSource)
                        End If

                        If Not _listBox.SelectedItems.Contains(clickedItem) _
                            AndAlso UIHelper.GetParentOfType(Of ScrollBar)(e.OriginalSource) Is Nothing _
                            AndAlso UIHelper.GetParentOfType(Of GridViewHeaderRowPresenter)(e.OriginalSource) Is Nothing Then
                            If TypeOf _listBox Is ListView AndAlso TypeOf CType(_listBox, ListView).View Is GridView Then
                                Dim hrp As GridViewHeaderRowPresenter = UIHelper.FindVisualChildren(Of GridViewHeaderRowPresenter)(_listBox)(0)
                                _headerHeight = hrp.ActualHeight
                            End If

                            _ptMin = New Point(_listBox.BorderThickness.Left, _listBox.BorderThickness.Top + _headerHeight)
                            _ptMax = New Point(_listBox.ActualWidth - _listBox.BorderThickness.Right - 1,
                                               _listBox.ActualHeight - _listBox.BorderThickness.Bottom - 1)
                            Dim vs As ScrollBar = _sv.Template.FindName("PART_VerticalScrollBar", _sv)
                            Dim hs As ScrollBar = _sv.Template.FindName("PART_HorizontalScrollBar", _sv)
                            If vs?.Visibility = Visibility.Visible Then _
                                _ptMax.X = _listBox.PointFromScreen(vs.PointToScreen(New Point(0, 0))).X - 1
                            If hs?.Visibility = Visibility.Visible Then _
                                _ptMax.Y = _listBox.PointFromScreen(hs.PointToScreen(New Point(0, 0))).Y - 1

                            _grid.Margin = New Thickness(_ptMin.X, _ptMin.Y, 0, 0)
                            _grid.Width = _ptMax.X - _ptMin.X + 1
                            _grid.Height = _ptMax.Y - _ptMin.Y + 1

                            _mouseDownPos = e.GetPosition(_listBox)
                            _canStartSelecting = True
                        Else
                            _canStartSelecting = False
                        End If
                    End If
                End Sub

            AddHandler _listBox.PreviewMouseUp,
                Sub(s As Object, e As MouseButtonEventArgs)
                    _canStartSelecting = False
                End Sub

            AddHandler _listBox.PreviewMouseMove,
                Sub(s As Object, e As MouseEventArgs)
                    Dim actualMousePos As Point = e.GetPosition(_listBox)

                    If Not Me.IsSelecting AndAlso _canStartSelecting _
                       AndAlso Not (TypeOf _mouseOriginalSourceDown Is TextBlock _
                                    AndAlso (CType(_mouseOriginalSourceDown, TextBlock).Tag = "PART_DisplayName" _
                                             OrElse CType(_mouseOriginalSourceDown, TextBlock).Name = "PART_DisplayName")) Then
                        If Math.Abs(actualMousePos.X - _mouseDownPos.X) > 2 OrElse Math.Abs(actualMousePos.Y - _mouseDownPos.Y) > 2 Then
                            Me.IsSelecting = True
                            _listBox.Focus()
                            _mouseDownPos.X += _sv.HorizontalOffset
                            _mouseDownPos.Y += _sv.VerticalOffset
                            _control.CaptureMouse()
                            _selectionRectangle.Margin = New Thickness(_mouseDownPos.X, _mouseDownPos.Y - _headerHeight, 0, 0)
                            _selectionRectangle.Width = 0
                            _selectionRectangle.Height = 0

                            _selectionRectangle.Visibility = Visibility.Visible
                            e.Handled = True
                        End If
                    End If
                End Sub

            AddHandler _control.PreviewMouseUp,
                Sub(s As Object, e As MouseButtonEventArgs)
                    If (Me.IsSelecting) Then
                        Me.IsSelecting = False
                        _canStartSelecting = False
                        _control.ReleaseMouseCapture()

                        _selectionRectangle.Visibility = Visibility.Collapsed

                        If Not _scrollTimer Is Nothing Then
                            _scrollTimer.Dispose()
                            _scrollTimer = Nothing
                        End If
                    End If
                End Sub

            AddHandler _control.PreviewMouseMove,
                Sub(s As Object, e As MouseEventArgs)
                    Dim actualMousePos As Point = e.GetPosition(_listBox)

                    If Me.IsSelecting Then
                        onAfterMouseMove()
                        e.Handled = True

                        If actualMousePos.X < _ptMin.X OrElse actualMousePos.Y < _ptMin.Y _
                            OrElse actualMousePos.X > _ptMax.X OrElse actualMousePos.Y > _ptMax.Y Then
                            If _scrollTimer Is Nothing Then
                                _scrollTimer = New Timer(
                                    Sub()
                                        UIHelper.OnUIThread(
                                            Sub()
                                                Dim scrollMousePos As Point = e.GetPosition(_listBox)
                                                If scrollMousePos.X < _ptMin.X Then
                                                    _sv.ScrollToHorizontalOffset(_sv.HorizontalOffset - 10)
                                                ElseIf scrollMousePos.X > _ptMax.X Then
                                                    _sv.ScrollToHorizontalOffset(_sv.HorizontalOffset + 10)
                                                End If
                                                If scrollMousePos.Y < _ptMin.Y Then
                                                    _sv.ScrollToVerticalOffset(_sv.VerticalOffset - 10)
                                                ElseIf scrollMousePos.Y > _ptMax.Y Then
                                                    _sv.ScrollToVerticalOffset(_sv.VerticalOffset + 10)
                                                End If

                                                onAfterMouseMove()
                                            End Sub)
                                    End Sub, Nothing, 30, 30)
                            End If
                        Else
                            If Not _scrollTimer Is Nothing Then
                                _scrollTimer.Dispose()
                                _scrollTimer = Nothing
                            End If
                        End If
                    End If
                End Sub
        End Sub

        Private Sub onAfterMouseMove()
            Dim mousePos As Point = Mouse.GetPosition(_listBox)
            If mousePos.X < _ptMin.X Then mousePos.X = _ptMin.X
            If mousePos.Y < _ptMin.Y Then mousePos.Y = _ptMin.Y
            If mousePos.X > _ptMax.X Then mousePos.X = _ptMax.X
            If mousePos.Y > _ptMax.Y Then mousePos.Y = _ptMax.Y
            mousePos.X += _sv.HorizontalOffset
            mousePos.Y += _sv.VerticalOffset

            If (_mouseDownPos.X < mousePos.X) Then
                _selectionRectangle.Margin = New Thickness(_mouseDownPos.X - _sv.HorizontalOffset, _selectionRectangle.Margin.Top, 0, 0)
                _selectionRectangle.Width = mousePos.X - _mouseDownPos.X
            Else
                _selectionRectangle.Margin = New Thickness(mousePos.X - _sv.HorizontalOffset, _selectionRectangle.Margin.Top, 0, 0)
                _selectionRectangle.Width = _mouseDownPos.X - mousePos.X
            End If

            If (_mouseDownPos.Y < mousePos.Y) Then
                _selectionRectangle.Margin = New Thickness(_selectionRectangle.Margin.Left, _mouseDownPos.Y - _headerHeight - _sv.VerticalOffset, 0, 0)
                _selectionRectangle.Height = mousePos.Y - _mouseDownPos.Y
            Else
                _selectionRectangle.Margin = New Thickness(_selectionRectangle.Margin.Left, mousePos.Y - _headerHeight - _sv.VerticalOffset, 0, 0)
                _selectionRectangle.Height = _mouseDownPos.Y - mousePos.Y
            End If

            Dim left As Double = _selectionRectangle.Margin.Left
            Dim top As Double = _selectionRectangle.Margin.Top + _headerHeight
            Dim width As Double = _selectionRectangle.Width
            Dim height As Double = _selectionRectangle.Height
            Dim rect As Rect = New Rect(left, top, width, height)

            For Each listViewItem In UIHelper.FindVisualChildren(Of ListViewItem)(_listBox)
                Dim bounds As Rect = listViewItem.TransformToAncestor(_listBox).TransformBounds(
                                New Rect(0.0, 0.0, listViewItem.ActualWidth, listViewItem.ActualHeight))
                If rect.IntersectsWith(bounds) AndAlso Not _listBox.SelectedItems.Contains(listViewItem.DataContext) Then
                    _listBox.SelectedItems.Add(listViewItem.DataContext)
                ElseIf Not rect.IntersectsWith(bounds) AndAlso _listBox.SelectedItems.Contains(listViewItem.DataContext) Then
                    _listBox.SelectedItems.Remove(listViewItem.DataContext)
                End If
            Next

            _listBox.Focus()
        End Sub

        Public Property IsSelecting As Boolean
            Get
                Return GetValue(IsSelectingProperty)
            End Get
            Set(ByVal value As Boolean)
                SetCurrentValue(IsSelectingProperty, value)
            End Set
        End Property
    End Class
End Namespace