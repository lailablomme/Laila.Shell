Imports System.Collections.ObjectModel
Imports System.Collections.Specialized
Imports System.ComponentModel
Imports System.Windows.Data

Namespace Helpers
    Public Class CustomObservableCollection(Of T)
        Inherits ObservableCollection(Of T)

        Private _lock As Object = New Object()

        Public Sub New()
            MyBase.New()

            BindingOperations.EnableCollectionSynchronization(Me, _lock)
        End Sub

        Public Sub AddRange(range As IEnumerable(Of T))
            Me.CheckReentrancy()

            For Each i In range
                Me.Items.Add(i)
            Next

            UIHelper.OnUIThread(
                Sub()
                    Me.OnCollectionChanged(New NotifyCollectionChangedEventArgs(NotifyCollectionChangedAction.Reset))
                End Sub)
        End Sub
    End Class
End Namespace