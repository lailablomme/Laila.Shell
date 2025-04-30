Imports System.Collections.Specialized
Imports System.ComponentModel
Imports System.Runtime.CompilerServices
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Data
Imports System.Windows.Input
Imports Laila.Shell.Helpers

Namespace Behaviors
    Public Class SharedSelectedItemsBehavior
        Public Shared ReadOnly SyncedSelectedItemsWithProperty As DependencyProperty =
            DependencyProperty.RegisterAttached(
                "SyncedSelectedItemsWith",
                GetType(Controls.BaseFolderView),
                GetType(SharedSelectedItemsBehavior),
                New PropertyMetadata(Nothing, AddressOf onSyncedSelectedItemsWithChanged))

        Public Shared Property ListViews As List(Of Tuple(Of ListView, Point)) = New List(Of Tuple(Of ListView, Point))()

        Private Shared _isWorking As Boolean
        Friend Shared Property IsSetInternally As Boolean

        Public Shared Sub SetSyncedSelectedItemsWith(element As DependencyObject, value As Controls.BaseFolderView)
            element.SetValue(SyncedSelectedItemsWithProperty, value)
        End Sub

        Public Shared Function GetSyncedSelectedItemsWith(element As DependencyObject) As Controls.BaseFolderView
            Return CType(element.GetValue(SyncedSelectedItemsWithProperty), Controls.BaseFolderView)
        End Function

        Public Shared Sub SetSyncedSelectedItems(element As DependencyObject, value As IEnumerable(Of Item))
            CType(element.GetValue(SyncedSelectedItemsWithProperty), Controls.BaseFolderView).SelectedItems = value
        End Sub

        Public Shared Function GetSyncedSelectedItems(element As DependencyObject) As IEnumerable(Of Item)
            Return CType(element.GetValue(SyncedSelectedItemsWithProperty), Controls.BaseFolderView).SelectedItems
        End Function

        Private Shared Sub onSyncedSelectedItemsWithChanged(d As ListView, e As DependencyPropertyChangedEventArgs)
            If Not ListViews.Exists(Function(i) i.Item1.Equals(d)) Then
                If Not ListViews.Exists(Function(i) i.Item1.Equals(d)) Then
                    ListViews.Add(New Tuple(Of ListView, Point)(d, New Point()))
                    RemoveHandler d.SelectionChanged, AddressOf listView_OnSelectionChanged
                    AddHandler d.SelectionChanged, AddressOf listView_OnSelectionChanged
                    RemoveHandler d.RequestBringIntoView, AddressOf listView_RequestBringIntoView
                    AddHandler d.RequestBringIntoView, AddressOf listView_RequestBringIntoView

                    Dim descriptor As DependencyPropertyDescriptor = DependencyPropertyDescriptor.FromProperty(Controls.BaseFolderView.SelectedItemsProperty, GetType(Controls.BaseFolderView))
                    If Not descriptor Is Nothing Then
                        descriptor.RemoveValueChanged(GetSyncedSelectedItemsWith(d), AddressOf baseFolderView_SelectedItemsChanged)
                        descriptor.AddValueChanged(GetSyncedSelectedItemsWith(d), AddressOf baseFolderView_SelectedItemsChanged)
                    End If

                    Dim sharedSelectedItems As IEnumerable(Of Item) = GetSyncedSelectedItems(d)
                    SelectItems(sharedSelectedItems)
                End If
            Else
                Dim sharedSelectedItems As IEnumerable(Of Item) = GetSyncedSelectedItems(d)
                SelectItems(sharedSelectedItems)
            End If
        End Sub

        Private Shared Sub baseFolderView_SelectedItemsChanged(sender As Controls.BaseFolderView, e2 As EventArgs)
            SelectItems(sender.SelectedItems)
        End Sub

        Private Shared Sub listView_RequestBringIntoView(sender As Object, e As RequestBringIntoViewEventArgs)
            e.Handled = True
        End Sub

        Public Shared Sub SelectItems(sharedSelectedItems As IEnumerable(Of Item))
            If _isWorking Then Return

            _isWorking = True
            If Not sharedSelectedItems Is Nothing Then
                For Each lv In ListViews
                    For Each item In lv.Item1.SelectedItems.Cast(Of Object).ToList()
                        If Not sharedSelectedItems.Contains(item) Then
                            lv.Item1.SelectedItems.Remove(item)
                        End If
                    Next
                    For Each item In sharedSelectedItems
                        If lv.Item1.Items.Contains(item) AndAlso Not lv.Item1.SelectedItems.Contains(item) Then
                            lv.Item1.SelectedItems.Add(item)
                        End If
                    Next
                Next
            Else
                For Each lv In ListViews
                    lv.Item1.SelectedItems.Clear()
                Next
            End If
            _isWorking = False
        End Sub

        Private Shared Sub listView_OnSelectionChanged(sender As Object, e As SelectionChangedEventArgs)
            If _isWorking Then Return
            If ListViews.Count = 0 Then Return

            If ListViews.Exists(Function(t) t.Item2.X = 0 And t.Item2.Y = 0) Then
                Dim parent As ListView = UIHelper.GetParentOfType(Of ListView)(sender)
                If Not parent Is Nothing Then
                    For Each lv In UIHelper.FindVisualChildren(Of ListView)(parent)
                        Dim t As Tuple(Of ListView, Point) = ListViews.FirstOrDefault(Function(i) i.Item1.Equals(lv))
                        ListViews.Remove(t)
                        ListViews.Add(New Tuple(Of ListView, Point)(lv, lv.TransformToAncestor(UIHelper.GetParentOfType(Of ListView)(lv)).Transform(New Point(0, 0))))
                    Next
                End If
                For Each item In ListViews.Where(Function(t) UIHelper.GetParentOfType(Of ListView)(t.Item1) Is Nothing).ToList()
                    ListViews.Remove(item)
                Next
            End If

            Dim result As List(Of Item) = New List(Of Item)()

            If Keyboard.Modifiers.HasFlag(ModifierKeys.Control) _
                        OrElse GetSyncedSelectedItemsWith(sender).IsSelecting _
                        OrElse IsSetInternally Then
                ' handle CTRL-click to select multiple items
                For Each lv In ListViews
                    For Each item In lv.Item1.SelectedItems
                        If Not result.Contains(item) Then
                            result.Add(item)
                        End If
                    Next
                Next
            ElseIf Keyboard.Modifiers.HasFlag(ModifierKeys.Shift) Then
                ' handle SHIFT-click to select multiple items
                Dim listViewsSorted As List(Of ListView) = ListViews _
                            .OrderBy(Function(t) t.Item2.Y) _
                            .Select(Function(t) t.Item1).ToList()
                Dim firstListView As ListView = listViewsSorted.FirstOrDefault(Function(lv) lv.SelectedItems.Count > 0)
                Dim lastListView As ListView = listViewsSorted.LastOrDefault(Function(lv) lv.SelectedItems.Count > 0)
                Dim el As IInputElement = Keyboard.FocusedElement
                If Not el Is Nothing Then
                    Dim clickedListViewItem As ListViewItem = Nothing
                    If TypeOf el Is ListViewItem Then
                        clickedListViewItem = el
                    Else
                        clickedListViewItem = UIHelper.GetParentOfType(Of ListViewItem)(el)
                    End If
                    If Not clickedListViewItem Is Nothing Then
                        If Not firstListView Is Nothing AndAlso Not lastListView Is Nothing Then
                            If firstListView.Equals(lastListView) Then
                                For Each item In firstListView.SelectedItems
                                    If Not result.Contains(item) Then
                                        result.Add(item)
                                    End If
                                Next
                            Else
                                Dim didSwap As Boolean
                                If firstListView.Equals(UIHelper.GetParentOfType(Of ListView)(clickedListViewItem)) Then
                                    Dim lvSwap As ListView = firstListView
                                    firstListView = lastListView
                                    lastListView = lvSwap
                                    didSwap = True
                                End If

                                Dim isSelecting As Boolean
                                For Each lv In listViewsSorted
                                    If lv.Equals(firstListView) Then
                                        Dim lastSelectedItemIndex As Integer = firstListView.Items.IndexOf(firstListView.SelectedItems(0))
                                        Dim list = If(didSwap, lv.Items.Cast(Of Item).Take(lastSelectedItemIndex + 1), lv.Items.Cast(Of Item).Skip(lastSelectedItemIndex))
                                        For Each item In list
                                            If Not result.Contains(item) Then
                                                result.Add(item)
                                            End If
                                        Next
                                        isSelecting = True
                                    ElseIf lv.Equals(lastListView) Then
                                        Dim lastSelectedItemIndex As Integer = lv.Items.IndexOf(clickedListViewItem.DataContext)
                                        Dim list = If(didSwap, lv.Items.Cast(Of Item).Skip(lastSelectedItemIndex), lv.Items.Cast(Of Item).Take(lastSelectedItemIndex + 1))
                                        For Each item In list
                                            If Not result.Contains(item) Then
                                                result.Add(item)
                                            End If
                                        Next
                                        isSelecting = False
                                    Else
                                        If isSelecting = Not didSwap Then
                                            result.AddRange(lv.Items.Cast(Of Item))
                                        End If
                                    End If
                                Next
                            End If
                        End If
                    End If
                End If
            Else
                ' handle normal click to select single item
                For Each item In sender.SelectedItems
                    If Not result.Contains(item) Then
                        result.Add(item)
                    End If
                Next
            End If

            SetSyncedSelectedItems(sender, result)
        End Sub
    End Class
End Namespace