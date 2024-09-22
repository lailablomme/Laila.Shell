Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Data
Imports System.Windows.Threading

Namespace Helpers
    Public Class SelectionHelper(Of TData)
        Private _control As ListBox
        Private _control2 As TreeView
        Private _isWorking As Boolean = False
        Private _selectedItem As Object

        Public SelectionChanged As Action =
             Sub()

             End Sub

        Public Sub New(control As ItemsControl)
            hook(control)
        End Sub

        Public Sub Unhook()
            If Not _control Is Nothing Then
                RemoveHandler _control.SelectionChanged, AddressOf listBox_SelectionChanged
            ElseIf Not _control2 Is Nothing Then
                RemoveHandler _control2.SelectedItemChanged, AddressOf treeView_SelectionChanged
            End If
        End Sub

        Private Sub hook(control As ItemsControl)
            If TypeOf control Is ListBox Then
                _control = control

                ' hook
                AddHandler _control.SelectionChanged, AddressOf listBox_SelectionChanged
            ElseIf TypeOf control Is TreeView Then
                _control2 = control

                ' hook
                AddHandler _control2.SelectedItemChanged, AddressOf treeView_SelectionChanged
            Else
                Throw New InvalidCastException()
            End If
        End Sub

        Private Sub treeView_SelectionChanged(sender As Object, e As RoutedPropertyChangedEventArgs(Of Object))
            If Not _isWorking Then
                ' notify changed
                SelectionChanged()
            End If
        End Sub

        Private Sub listBox_SelectionChanged(sender As Object, e As SelectionChangedEventArgs)
            If Not _isWorking Then
                ' notify changed
                SelectionChanged()
            End If
        End Sub

        Private Sub scrollIntoView(lb As ListBox, item As Object)
            ' somehow, when scrolling an item into view in one listbox, all the other listboxes in the same context
            ' get a layoutupdated call, resetting their scroll offsets. so here we save them first, scroll the item into
            ' view and then set them back. example: midden->gereedschappen->prijzen

            Dim p As DependencyObject = lb
            While Not UIHelper.GetParent(p) Is Nothing
                p = UIHelper.GetParent(p)
            End While

            Dim d As Dictionary(Of ScrollViewer, Tuple(Of Double, Double)) = New Dictionary(Of ScrollViewer, Tuple(Of Double, Double))()

            For Each sv In UIHelper.FindVisualChildren(Of ScrollViewer)(p)
                If Not lb.Equals(UIHelper.GetParentOfType(Of ListBox)(sv)) Then
                    d.Add(sv, New Tuple(Of Double, Double)(sv.HorizontalOffset, sv.VerticalOffset))
                End If
            Next

            lb.ScrollIntoView(item)

            For Each i In d
                i.Key.ScrollToHorizontalOffset(i.Value.Item1)
                i.Key.ScrollToVerticalOffset(i.Value.Item2)
            Next
        End Sub

        Public ReadOnly Property SelectedItems As IEnumerable(Of TData)
            Get
                If Not _control Is Nothing Then
                    Dim result As IEnumerable(Of TData) = Nothing
                    If _control.SelectionMode = SelectionMode.Single Then
                        If _control.SelectedItem Is Nothing Then
                            result = {}
                        Else
                            result = {_control.SelectedItem}
                        End If
                    Else
                        result = _control.SelectedItems.Cast(Of TData)()
                    End If
                    Return result
                ElseIf Not _control2 Is Nothing Then
                    Dim result As IEnumerable(Of TData) = Nothing
                    If _control2.SelectedItem Is Nothing Then
                        result = {}
                    Else
                        result = {_control2.SelectedItem}
                    End If
                    Return result
                Else
                    Throw New InvalidOperationException()
                End If
            End Get
        End Property
        Public Sub SetSelectedItems(value As IEnumerable(Of TData))
            If Not _control Is Nothing Then
                ' Wait for any databinding to finish
                If Not Application.Current.Dispatcher.CheckAccess() Then
                    Application.Current.Dispatcher.Invoke(
                        Sub()
                        End Sub, Threading.DispatcherPriority.DataBind)
                End If

                System.Windows.Application.Current.Dispatcher.Invoke(
                     Sub()
                         ' clean
                         Dim selectedItems As IEnumerable(Of TData) = value.Where(Function(v) Not v Is Nothing)

                         ' turn off notifications  
                         _isWorking = True

                         If _control.SelectionMode = SelectionMode.Single Then
                             If selectedItems.Count = 0 Then
                                 If Not _control.SelectedItem Is Nothing Then
                                     _control.SelectedItem = Nothing
                                     SelectionChanged()
                                 End If
                             ElseIf selectedItems.Count = 1 Then
                                 If Not selectedItems(0).Equals(_control.SelectedItem) Then
                                     _control.SelectedItem = selectedItems(0)
                                     '_control.ScrollIntoView(selectedItems(0))
                                     scrollIntoView(_control, selectedItems(0))
                                     SelectionChanged()
                                 End If
                             Else
                                 Throw New ArgumentException("You cannot select multiple items when SelectionMode is Single.")
                             End If
                         Else
                             Dim isSame As Boolean = True
                             If _control.SelectedItems.Count <> selectedItems.Count Then
                                 isSame = False
                             Else
                                 For Each i In _control.SelectedItems
                                     If Not selectedItems.Contains(i) Then
                                         isSame = False
                                         Exit For
                                     End If
                                 Next

                                 If isSame Then
                                     For Each i In selectedItems
                                         If Not _control.SelectedItems.Contains(i) Then
                                             isSame = False
                                             Exit For
                                         End If
                                     Next
                                 End If
                             End If

                             If Not isSame Then
                                 ' add selected items
                                 _control.SelectedItems.Clear()
                                 For Each item In selectedItems
                                     _control.SelectedItems.Add(_control.Items.Cast(Of TData).FirstOrDefault(Function(i) i.Equals(item)))
                                 Next
                                 If selectedItems.Count > 0 Then
                                     '_control.ScrollIntoView(selectedItems(0))
                                     scrollIntoView(_control, selectedItems(0))
                                 End If

                                 ' notify change
                                 SelectionChanged()
                             End If
                         End If

                         ' turn back on notifications  
                         _isWorking = False
                     End Sub)
            ElseIf Not _control2 Is Nothing Then
                System.Windows.Application.Current.Dispatcher.Invoke(
                    Sub()
                        ' clean
                        Dim selectedItems As IEnumerable(Of TData) = value.Where(Function(v) Not v Is Nothing)

                        ' turn off notifications  
                        _isWorking = True

                        If selectedItems.Count = 0 Then
                            _selectedItem = Nothing
                            If Not _control2.SelectedItem Is Nothing Then
                                Dim tvi As TreeViewItem = findTVI(_control2, _control2.SelectedItem)
                                If Not tvi Is Nothing Then
                                    tvi.IsSelected = False
                                    ' notify change
                                    SelectionChanged()
                                End If
                            End If
                        ElseIf selectedItems.Count = 1 Then
                            _selectedItem = selectedItems(0)
                            If Not selectedItems(0).Equals(_control2.SelectedItem) Then
                                Dim tvi As TreeViewItem = findTVI(_control2, selectedItems(0))
                                If Not tvi Is Nothing Then
                                    tvi.IsSelected = True
                                    tvi.BringIntoView()
                                    ' notify change
                                    SelectionChanged()
                                End If
                            End If
                        Else
                            Throw New ArgumentException("You cannot select multiple items when SelectionMode is Single.")
                        End If

                        ' turn back on notifications  
                        _isWorking = False
                    End Sub, DispatcherPriority.Loaded)
            Else
                Throw New InvalidOperationException()
            End If
        End Sub

        Private Function findTVI(ic As ItemsControl, item As Object) As TreeViewItem
            Dim i As Integer = 0
            For Each item2 In ic.Items
                Dim t As TreeViewItem = ic.ItemContainerGenerator.ContainerFromIndex(i)
                If Not t Is Nothing AndAlso t.DataContext.Equals(item) Then
                    Return t
                ElseIf Not t Is Nothing Then
                    t = findTVI(t, item)
                    If Not t Is Nothing Then
                        Return t
                    End If
                End If
                i += 1
            Next
            'Dim found As TreeViewItem = ic.ItemContainerGenerator.ContainerFromItem(item)
            'If Not found Is Nothing Then
            '    Return found
            'Else
            '    For Each i In ic.Items
            '        Dim itvi As TreeViewItem = ic.ItemContainerGenerator.ContainerFromItem(i)
            '        If Not itvi Is Nothing Then
            '            found = findTVI(itvi, item)
            '            If Not found Is Nothing Then
            '                Return found
            '            End If
            '        End If
            '    Next
            'End If
            Return Nothing
        End Function
    End Class
End Namespace