Imports System.Collections.ObjectModel
Imports System.Collections.Specialized
Imports System.ComponentModel
Imports System.Windows.Data

Namespace Helpers
    Public Class CustomObservableCollection(Of T)
        Inherits ObservableCollection(Of T)

        Public Sub UpdateRange(itemsToAdd As IEnumerable(Of T), itemsToRemove As IEnumerable(Of T))
            Me.CheckReentrancy()

            If Not itemsToAdd Is Nothing Then
                For Each i In itemsToAdd
                    Me.Items.Add(i)
                Next
            End If

            If Not itemsToRemove Is Nothing Then
                For Each i In itemsToRemove
                    Me.Items.Remove(i)
                Next
            End If

            Me.OnCollectionChanged(New NotifyCollectionChangedEventArgs(NotifyCollectionChangedAction.Reset))
        End Sub
    End Class
End Namespace