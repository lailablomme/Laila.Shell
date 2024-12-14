Imports System.Collections.ObjectModel
Imports System.Collections.Specialized
Imports System.ComponentModel
Imports System.Windows.Data

Namespace Helpers
    Public Class CustomObservableCollection(Of T)
        Inherits ObservableCollection(Of T)

        Public Sub AddRange(range As IEnumerable(Of T))
            Me.CheckReentrancy()

            For Each i In range
                Me.Items.Add(i)
            Next

            Me.OnCollectionChanged(New NotifyCollectionChangedEventArgs(NotifyCollectionChangedAction.Reset))
        End Sub
    End Class
End Namespace