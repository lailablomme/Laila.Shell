﻿Imports System.Collections.ObjectModel
Imports System.Collections.Specialized
Imports System.ComponentModel
Imports System.Windows.Data

Namespace Helpers
    Public Class ItemsCollection(Of T)
        Inherits ObservableCollection(Of T)

        Public Property Lock As Object = New Object()

        Public Sub RemoveWithoutNotifying(item As T)
            SyncLock Me.Lock
                Me.Items.Remove(item)
            End SyncLock
        End Sub

        Public Sub UpdateRange(itemsToAdd As IEnumerable(Of T), itemsToRemove As IEnumerable(Of T))
            Me.CheckReentrancy()

            SyncLock Me.Lock
                If Not itemsToRemove Is Nothing Then
                    For Each i In itemsToRemove
                        Me.Items.Remove(i)
                    Next
                End If

                If Not itemsToAdd Is Nothing Then
                    For Each i In itemsToAdd
                        Me.Items.Add(i)
                    Next
                End If
            End SyncLock

            If If(itemsToAdd?.Count > 0, False) OrElse If(itemsToRemove?.Count > 0, False) Then
                Me.OnCollectionChanged(New NotifyCollectionChangedEventArgs(NotifyCollectionChangedAction.Reset))
            End If
        End Sub

        Public Sub InsertSorted(newItem As T, comparer As IComparer)
            SyncLock Me.Lock
                ' Find the correct index using the comparer
                Dim index As Integer = FindInsertionIndex(newItem, comparer)

                ' Insert the item at the correct position
                If index >= 0 AndAlso index <= Me.Items.Count Then
                    Me.Items.Insert(index, newItem)
                    Me.OnCollectionChanged(New NotifyCollectionChangedEventArgs(NotifyCollectionChangedAction.Add, newItem, index))
                End If
            End SyncLock
        End Sub

        ' Find the correct index for insertion using the IComparer
        Private Function FindInsertionIndex(item As T, comparer As IComparer) As Integer
            Dim low As Integer = 0
            Dim high As Integer = Me.Items.Count - 1

            While low <= high
                Dim mid As Integer = (low + high) \ 2
                Dim comparisonResult As Integer = comparer.Compare(Me.Items(mid), item)

                If comparisonResult = 0 Then
                    Return mid ' Item already exists, return that index
                ElseIf comparisonResult < 0 Then
                    low = mid + 1
                Else
                    high = mid - 1
                End If
            End While

            Return low ' Return the insertion index where the item should be inserted
        End Function

        Protected Overrides Sub OnCollectionChanged(e As NotifyCollectionChangedEventArgs)
            UIHelper.OnUIThread(
                Sub()
                    MyBase.OnCollectionChanged(e)
                End Sub)
        End Sub

        Protected Overrides Sub OnPropertyChanged(e As PropertyChangedEventArgs)
            UIHelper.OnUIThread(
                Sub()
                    MyBase.OnPropertyChanged(e)
                End Sub)
        End Sub
    End Class
End Namespace