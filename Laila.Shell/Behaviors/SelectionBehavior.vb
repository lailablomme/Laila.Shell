Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Input
Imports System.Windows.Media
Imports System.Windows.Shapes
Imports Laila.Shell.Helpers
Imports Microsoft.Xaml.Behaviors

Namespace Behaviors
    Public Class SelectionBehavior
        Inherits Behavior(Of ListView)

        Private _listView As ListView
        Private _selectionRectangle As Rectangle
        Private _isSelecting As Boolean
        Private _mouseDownPos As Point
        Private _control As Control
        Private _sv As ScrollViewer
        Private _headerHeight As Double

        Protected Overrides Sub OnAttached()
            MyBase.OnAttached()

            _listView = Me.AssociatedObject
            _control = New Control() With {.Opacity = 0}

            AddHandler _listView.Loaded,
                Sub(s As Object, e As EventArgs)
                    _sv = UIHelper.FindVisualChildren(Of ScrollViewer)(_listView)(0)
                    _selectionRectangle = New Rectangle() With {
                        .Stroke = Brushes.SkyBlue,
                        .Fill = New SolidColorBrush(Color.FromArgb(75, Colors.SkyBlue.R, Colors.SkyBlue.G, Colors.SkyBlue.B)),
                        .Visibility = Visibility.Collapsed,
                        .HorizontalAlignment = HorizontalAlignment.Left,
                        .VerticalAlignment = VerticalAlignment.Top
                    }
                    Dim grid As Grid = New Grid()
                    Dim content As Object = _sv.Content
                    _sv.Content = grid
                    grid.Children.Add(content)
                    grid.Children.Add(_selectionRectangle)
                    grid.Children.Add(_control)
                End Sub

            AddHandler _listView.PreviewMouseDown,
                Sub(s As Object, e As MouseButtonEventArgs)
                    If e.LeftButton = Input.MouseButtonState.Pressed AndAlso Not e.OriginalSource Is Nothing Then
                        Dim listViewItem As ListViewItem = UIHelper.GetParentOfType(Of ListViewItem)(e.OriginalSource)
                        Dim clickedItem As Item = listViewItem?.DataContext

                        If Not _listView.SelectedItems.Contains(clickedItem) Then
                            Dim hrp As GridViewHeaderRowPresenter = UIHelper.FindVisualChildren(Of GridViewHeaderRowPresenter)(_listView)(0)
                            hrp.Measure(New Size(_listView.ActualHeight, _listView.ActualHeight))
                            _headerHeight = hrp.DesiredSize.Height
                            _mouseDownPos = e.GetPosition(_listView)
                            If _mouseDownPos.Y > _headerHeight Then
                                _isSelecting = True
                                _listView.Focus()
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
                    End If
                End Sub

            AddHandler _control.PreviewMouseUp,
                Sub(s As Object, e As MouseButtonEventArgs)
                    If (_isSelecting) Then
                        _isSelecting = False
                        _control.ReleaseMouseCapture()

                        _selectionRectangle.Visibility = Visibility.Collapsed
                    End If
                End Sub

            AddHandler _control.PreviewMouseMove,
                Sub(s As Object, e As MouseEventArgs)
                    If (_isSelecting) Then
                        Dim mousePos As Point = e.GetPosition(_listView)
                        Dim actualMousePos As Point = e.GetPosition(_listView)
                        mousePos.X += _sv.HorizontalOffset
                        mousePos.Y += _sv.VerticalOffset

                        If (_mouseDownPos.X < mousePos.X) Then
                            _selectionRectangle.Margin = New Thickness(_mouseDownPos.X, _selectionRectangle.Margin.Top, 0, 0)
                            _selectionRectangle.Width = mousePos.X - _mouseDownPos.X
                        Else
                            _selectionRectangle.Margin = New Thickness(mousePos.X, _selectionRectangle.Margin.Top, 0, 0)
                            _selectionRectangle.Width = _mouseDownPos.X - mousePos.X
                        End If

                        If (_mouseDownPos.Y < mousePos.Y) Then
                            _selectionRectangle.Margin = New Thickness(_selectionRectangle.Margin.Left, _mouseDownPos.Y - _headerHeight, 0, 0)
                            _selectionRectangle.Height = mousePos.Y - _mouseDownPos.Y
                        Else
                            _selectionRectangle.Margin = New Thickness(_selectionRectangle.Margin.Left, mousePos.Y - _headerHeight, 0, 0)
                            _selectionRectangle.Height = _mouseDownPos.Y - mousePos.Y
                        End If

                        Dim left As Double, top As Double
                        Dim width As Double = _selectionRectangle.Width, height As Double = _selectionRectangle.Height
                        If _selectionRectangle.Margin.Left - _sv.HorizontalOffset < 0 Then
                            Dim diff As Double = -(_selectionRectangle.Margin.Left - _sv.HorizontalOffset)
                            left = 0 : width -= diff
                        Else
                            left = _selectionRectangle.Margin.Left - _sv.HorizontalOffset
                        End If
                        If _selectionRectangle.Margin.Top - _sv.VerticalOffset + _headerHeight < 0 Then
                            Dim diff As Double = -(_selectionRectangle.Margin.Top - _sv.VerticalOffset + _headerHeight)
                            top = 0 : height -= diff
                        Else
                            top = _selectionRectangle.Margin.Top - _sv.VerticalOffset + _headerHeight
                        End If
                        If left + width > _listView.ActualWidth Then
                            width = _listView.ActualWidth - left
                        End If
                        If top + height > _listView.ActualHeight Then
                            height = _listView.ActualHeight - top
                        End If
                        Dim rect As Rect = New Rect(left, top, width, height)

                        For Each item In _listView.Items
                            Dim listViewItem As ListViewItem = _listView.ItemContainerGenerator.ContainerFromItem(item)
                            If Not listViewItem Is Nothing Then
                                Dim bounds As Rect = listViewItem.TransformToAncestor(_listView).TransformBounds(
                                    New Rect(0.0, 0.0, listViewItem.ActualWidth, listViewItem.ActualHeight))
                                If rect.IntersectsWith(bounds) AndAlso Not _listView.SelectedItems.Contains(item) Then
                                    _listView.SelectedItems.Add(item)
                                ElseIf Not rect.IntersectsWith(bounds) AndAlso _listView.SelectedItems.Contains(item) Then
                                    _listView.SelectedItems.Remove(item)
                                End If
                            End If
                        Next

                        e.Handled = True
                    End If
                End Sub
        End Sub

        Public ReadOnly Property IsSelecting As Boolean
            Get
                Return _isSelecting
            End Get
        End Property
    End Class
End Namespace