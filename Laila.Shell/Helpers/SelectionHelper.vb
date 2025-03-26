Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Data
Imports System.Windows.Input

Namespace Helpers
    Public Class SelectionHelper(Of TData)
        Private _control As ListBox
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
            End If
        End Sub

        Private Sub hook(control As ItemsControl)
            If TypeOf control Is ListBox Then
                _control = control

                ' hook
                AddHandler _control.SelectionChanged, AddressOf listBox_SelectionChanged
            Else
                Throw New InvalidCastException()
            End If
        End Sub

        Private Sub listBox_SelectionChanged(sender As Object, e As SelectionChangedEventArgs)
            If Not _isWorking Then
                _isWorking = True
                ' notify changed
                SelectionChanged()
                _isWorking = False
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
                Else
                    Throw New InvalidOperationException()
                End If
            End Get
        End Property

        Public Sub SetSelectedItems(value As IEnumerable(Of TData), Optional doScrollIntoView As Boolean = True)
            If Not _control Is Nothing Then
                If Not _isWorking Then
                    _isWorking = True
                    ' clean
                    Dim selectedItems As IEnumerable(Of TData) = If(value Is Nothing, {}, value.Where(Function(v) Not v Is Nothing)).ToList()

                    If _control.SelectionMode = SelectionMode.Single Then
                        If selectedItems.Count = 0 Then
                            If Not _control.SelectedItem Is Nothing Then
                                _control.SelectedItem = Nothing
                                SelectionChanged()
                            End If
                        ElseIf selectedItems.Count = 1 Then
                            Dim isChanged As Boolean = False
                            If Not selectedItems(0).Equals(_control.SelectedItem) Then
                                _control.SelectedItem = selectedItems(0)
                                isChanged = True
                            End If
                            If doScrollIntoView Then
                                scrollIntoView(_control, selectedItems(0))
                            End If
                            If isChanged Then
                                SelectionChanged()
                            End If
                        Else
                            Throw New ArgumentException("You cannot select multiple items when SelectionMode is Single.")
                        End If
                    Else
                        ' add selected items
                        If _control.Items.Count = selectedItems.Count Then
                            _control.SelectAll()
                        Else
                            _control.SelectedItems.Clear()
                            For Each item In selectedItems
                                _control.SelectedItems.Add(item)
                            Next
                        End If
                        ' scroll into view?
                        If doScrollIntoView AndAlso selectedItems.Count > 0 Then
                            scrollIntoView(_control, selectedItems(0))
                        End If

                        ' notify change
                        SelectionChanged()
                    End If

                    ' turn back on notifications  
                    _isWorking = False
                End If
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
            Return Nothing
        End Function
    End Class
End Namespace