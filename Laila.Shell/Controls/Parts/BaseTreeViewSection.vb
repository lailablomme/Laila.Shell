Imports System.Collections.ObjectModel
Imports System.Collections.Specialized
Imports Laila.Shell.Helpers

Namespace Controls.Parts
    Public Class BaseTreeViewSection
        Inherits NotifyPropertyChangedBase
        Implements IDisposable

        Private _items As ObservableCollection(Of Item) = New ObservableCollection(Of Item)()
        Private _treeView As TreeView
        Protected disposedValue As Boolean

        Public Sub New()
            AddHandler Me.Items.CollectionChanged, AddressOf items_CollectionChanged
        End Sub

        Private Sub items_CollectionChanged(s As Object, e As NotifyCollectionChangedEventArgs)
            Select Case e.Action
                Case NotifyCollectionChangedAction.Add
                    For Each item In e.NewItems
                        Dim i As Item = item
                        i.TreeViewSection = Me
                    Next
                Case NotifyCollectionChangedAction.Remove
                    For Each item In e.OldItems
                        Dim i As Item = item
                        i.TreeViewSection = Nothing
                    Next
                Case NotifyCollectionChangedAction.Reset
                    Throw New NotSupportedException()
            End Select
        End Sub

        Public ReadOnly Property Items As ObservableCollection(Of Item)
            Get
                Return _items
            End Get
        End Property

        Friend Overridable Sub Initialize()

        End Sub

        Friend Property TreeView As TreeView
            Get
                Return _treeView
            End Get
            Set(value As TreeView)
                SetValue(_treeView, value)
            End Set
        End Property

        Protected Overridable Sub Dispose(disposing As Boolean)
            If Not disposedValue Then
                If disposing Then
                    ' dispose managed state (managed objects)
                End If

                ' free unmanaged resources (unmanaged objects) and override finalizer
                ' set large fields to null
                disposedValue = True
            End If
        End Sub

        Public Sub Dispose() Implements IDisposable.Dispose
            ' Do not change this code. Put cleanup code in 'Dispose(disposing As Boolean)' method
            Dispose(disposing:=True)
            GC.SuppressFinalize(Me)
        End Sub
    End Class
End Namespace