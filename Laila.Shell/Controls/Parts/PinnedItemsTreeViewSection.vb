Imports System.Collections.ObjectModel
Imports System.Runtime.InteropServices.ComTypes
Imports Laila.Shell.Events
Imports Laila.Shell.Helpers
Imports Laila.Shell.Interfaces
Imports Laila.Shell.Interop
Imports Laila.Shell.Interop.DragDrop
Imports Laila.Shell.Interop.Items

Namespace Controls.Parts
    Public Class PinnedItemsTreeViewSection
        Inherits BaseTreeViewSection
        Implements ISupportDragInsert

        Friend Overrides Sub Initialize()
            updatePinnedItems()

            AddHandler PinnedItems.ItemPinned,
                Sub(s As Object, e As PinnedItemEventArgs)
                    updatePinnedItems()
                End Sub
            AddHandler PinnedItems.ItemUnpinned,
                Sub(s As Object, e As PinnedItemEventArgs)
                    updatePinnedItems()
                End Sub

            AddHandler Shell.Notification,
                Sub(s As Object, e As NotificationEventArgs)
                    Select Case e.Event
                        Case SHCNE.RMDIR, SHCNE.DELETE
                            UIHelper.OnUIThread(
                                Sub()
                                    If Not Me.Items.FirstOrDefault(Function(i) _
                                        Not i.disposedValue _
                                        AndAlso Not i.FullPath Is Nothing _
                                        AndAlso i.FullPath.Equals(e.Item1.FullPath)) Is Nothing Then
                                        updatePinnedItems()
                                    End If
                                End Sub)
                        Case SHCNE.UPDATEDIR
                            UIHelper.OnUIThread(
                                Sub()
                                    If (Not Me.Items.FirstOrDefault(
                                        Function(i)
                                            If Not i.disposedValue Then
                                                Return (Not i.Parent Is Nothing _
                                                    AndAlso Not i.Parent.FullPath Is Nothing _
                                                    AndAlso i.Parent.FullPath.Equals(e.Item1.FullPath))
                                            Else
                                                Return False
                                            End If
                                        End Function) Is Nothing _
                                    OrElse Shell.Desktop.FullPath.Equals(e.Item1.FullPath)) Then
                                        updatePinnedItems()
                                    End If
                                End Sub)
                    End Select
                End Sub
        End Sub

        Private Sub updatePinnedItems()
            Dim selectedPidl As Pidl = Nothing
            If Me.Items.Contains(Me.TreeView.SelectedItem) Then
                selectedPidl = Me.TreeView.SelectedItem.Pidl.Clone()
            End If

            For Each item In Me.Items.ToList()
                Me.Items.Remove(item)
            Next

            Me.Items.Add(New SeparatorFolder())

            For Each item In PinnedItems.GetPinnedItems()
                Me.Items.Add(item)
                If item.Pidl.Equals(selectedPidl) Then Me.TreeView.SetSelectedItem(item)
            Next

            If Not selectedPidl Is Nothing Then
                selectedPidl.Dispose()
            End If

            If Me.Items.Count = 1 Then
                Me.Items.Add(New PinnedItemsPlaceholderFolder())
            End If
        End Sub

        Private ReadOnly Property ISupportDragInsert_Items As ObservableCollection(Of Item) Implements ISupportDragInsert.Items
            Get
                Return Me.Items
            End Get
        End Property

        Public Function DragInsertBefore(dataObject As IDataObject, files As List(Of Item), index As Integer) As Interop.HRESULT Implements ISupportDragInsert.DragInsertBefore
            Dim canPinItem As Boolean =
                (Me.Items.Count = 2 AndAlso TypeOf Me.Items(1) Is PinnedItemsPlaceholderFolder AndAlso index = -1) _
                OrElse index = 0 _
                OrElse (index + 1 > Me.Items.Count - 1 AndAlso Me.Items(Me.Items.Count - 1).IsPinned) _
                OrElse Me.Items(index + 1).IsPinned
            If canPinItem Then
                WpfDragTargetProxy.SetDropDescription(dataObject, DROPIMAGETYPE.DROPIMAGE_LINK, "Pin to %1", "Quick access")
                Return HRESULT.S_OK
            Else
                Return HRESULT.S_FALSE
            End If
        End Function

        Public Function Drop(dataObject As IDataObject, files As List(Of Item), index As Integer) As Interop.HRESULT Implements ISupportDragInsert.Drop
            Try
                PinnedItems.IsNotifying = False

                For Each file In files
                    Dim unpinnedIndex As Integer = PinnedItems.UnpinItem(file.Pidl)
                    If unpinnedIndex <> -1 AndAlso unpinnedIndex < index Then
                        If index <> 0 Then index -= 1
                    End If
                Next
                For Each file In files
                    PinnedItems.PinItem(file, index)
                    If index <> -1 Then index += 1
                Next

                updatePinnedItems()

                Return HRESULT.S_OK
            Finally
                PinnedItems.IsNotifying = True
                PinnedItems.NotifyReset()
            End Try
        End Function
    End Class
End Namespace