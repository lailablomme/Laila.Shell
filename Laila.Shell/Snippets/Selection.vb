'Private _isMouseDown As Boolean
'Private _mouseDownPos As Point

'Private Sub ListBox_MouseDown(sender As Object, e As Input.MouseButtonEventArgs)
'    If e.Source.Equals(listBox) AndAlso e.LeftButton = Input.MouseButtonState.Pressed Then
'        listBox.SetValue(GongSolutions.Wpf.DragDrop.DragDrop.IsDragSourceProperty, False)

'        _isMouseDown = True
'        _mouseDownPos = e.GetPosition(listBox)
'        listBox.CaptureMouse()

'        Canvas.SetLeft(selectionBox, _mouseDownPos.X)
'        Canvas.SetTop(selectionBox, _mouseDownPos.Y)
'        selectionBox.Width = 0
'        selectionBox.Height = 0

'        selectionBox.Visibility = Visibility.Visible

'        e.Handled = True
'    End If
'End Sub

'Private Sub ListBox_MouseUp(sender As Object, e As Input.MouseButtonEventArgs)
'    If (_isMouseDown) Then
'        listBox.SetValue(GongSolutions.Wpf.DragDrop.DragDrop.IsDragSourceProperty, True)

'        _isMouseDown = False
'        listBox.ReleaseMouseCapture()

'        selectionBox.Visibility = Visibility.Collapsed
'    End If
'End Sub

'Private Sub ListBox_MouseMove(sender As Object, e As Input.MouseEventArgs)
'    If (_isMouseDown) Then
'        Dim mousePos As Point = e.GetPosition(listBox)

'        If (_mouseDownPos.X < mousePos.X) Then
'            Canvas.SetLeft(selectionBox, _mouseDownPos.X)
'            selectionBox.Width = mousePos.X - _mouseDownPos.X
'        Else
'            Canvas.SetLeft(selectionBox, mousePos.X)
'            selectionBox.Width = _mouseDownPos.X - mousePos.X
'        End If

'        If (_mouseDownPos.Y < mousePos.Y) Then
'            Canvas.SetTop(selectionBox, _mouseDownPos.Y)
'            selectionBox.Height = mousePos.Y - _mouseDownPos.Y
'        Else
'            Canvas.SetTop(selectionBox, mousePos.Y)
'            selectionBox.Height = _mouseDownPos.Y - mousePos.Y
'        End If

'        Dim rect As Rect = New Rect(Canvas.GetLeft(selectionBox), Canvas.GetTop(selectionBox), selectionBox.Width, selectionBox.Height)

'        listBox.SelectedItems.Clear()
'        For Each item In listBox.Items
'            Dim listBoxItem As ListBoxItem = listBox.ItemContainerGenerator.ContainerFromItem(item)
'            Dim bounds As Rect = listBoxItem.TransformToAncestor(listBox).TransformBounds(New Rect(0.0, 0.0, listBoxItem.ActualWidth, listBoxItem.ActualHeight))
'            If rect.IntersectsWith(bounds) Then
'                listBox.SelectedItems.Add(item)
'            End If
'        Next
'    End If
'End Sub