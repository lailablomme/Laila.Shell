﻿Imports System.IO
Imports Laila.Shell.Events
Imports Laila.Shell.FrequentFolders
Imports LiteDB

Public Class PinnedItems
    Public Shared Event ItemPinned(sender As Object, e As PinnedItemEventArgs)
    Public Shared Event ItemUnpinned(sender As Object, e As PinnedItemEventArgs)

    Public Shared Function GetPinnedItems() As IEnumerable(Of Item)
        Dim pinnedItems As List(Of PinnedItem)

        Using db = New LiteDatabase(getDBFileName())
            Dim collection As ILiteCollection(Of PinnedItem) = db.GetCollection(Of PinnedItem)("PinnedItems")

            pinnedItems = collection.Query() _
                .OrderBy(Function(f) f.Index).ToList()
        End Using

        ' return existing pinned items, delete others
        Dim tcs As New TaskCompletionSource(Of IEnumerable(Of Item))

        Shell.STATaskQueue.Add(
            Sub()
                Try
                    Dim existingPinnedItems As List(Of Item) = New List(Of Item)()
                    For Each pinnedItem In pinnedItems
                        Dim pidl As Pidl
                        Try
                            pidl = New Pidl(pinnedItem.Pidl)
                            Dim i As Item = Item.FromPidl(pidl.AbsolutePIDL, Nothing)
                            If Not i Is Nothing Then
                                existingPinnedItems.Add(i)
                            Else
                                UnpinItem(pidl)
                            End If
                        Finally
                            If Not pidl Is Nothing Then
                                pidl.Dispose()
                            End If
                        End Try
                    Next

                    tcs.SetResult(existingPinnedItems)
                Catch ex As Exception
                    tcs.SetException(ex)
                End Try
            End Sub)

        tcs.Task.Wait(Shell.ShuttingDownToken)
        If Not Shell.ShuttingDownToken.IsCancellationRequested Then
            Return tcs.Task.Result
        Else
            Return {}
        End If
    End Function

    Public Shared Function GetIsPinned(item As Item) As Boolean
        Using db = New LiteDatabase(getDBFileName())
            ' register in db
            Dim collection As ILiteCollection(Of PinnedItem) = db.GetCollection(Of PinnedItem)("PinnedItems")
            Dim pinnedItems As IEnumerable(Of PinnedItem) = collection.Find(Function(f) f.Pidl.Equals(item.Pidl.ToString()))
            Return pinnedItems.Count = 1
        End Using
    End Function

    Public Shared Sub PinItem(item As Item, Optional newIndex As Long = -1)
        Dim e As PinnedItemEventArgs

        Using db = New LiteDatabase(getDBFileName())
            ' register in db
            Dim collection As ILiteCollection(Of PinnedItem) = db.GetCollection(Of PinnedItem)("PinnedItems")
            Dim pinnedItems As IEnumerable(Of PinnedItem) = collection.Find(Function(f) f.Pidl.Equals(item.Pidl.ToString()))
            If pinnedItems.Count = 0 Then
                If newIndex >= 0 Then
                    ' move items down
                    For Each movePinnedItem In collection.Query().Where(Function(i) i.Index >= newIndex).ToList()
                        movePinnedItem.Index += 1
                        collection.Update(movePinnedItem)
                    Next
                End If

                ' insert new item
                Dim pinnedItem As PinnedItem = New PinnedItem() With {
                    .Pidl = item.Pidl.ToString(),
                    .Index = If(newIndex >= 0, newIndex, If(collection.Query().Count > 0,
                        collection.Query().OrderByDescending(Function(i) i.Index).First().Index + 1, 0))
                }
                collection.Insert(pinnedItem)
                collection.EnsureIndex(Of String)(Function(f) f.Pidl)
                e = New PinnedItemEventArgs() With {.Pidl = item.Pidl, .Index = pinnedItem.Index}
            End If
        End Using

        If Not e Is Nothing Then RaiseEvent ItemPinned(Nothing, e)
    End Sub

    Public Shared Function UnpinItem(pidl As Pidl) As Integer
        Dim e As PinnedItemEventArgs

        Using db = New LiteDatabase(getDBFileName())
            ' register in db
            Dim collection As ILiteCollection(Of PinnedItem) = db.GetCollection(Of PinnedItem)("PinnedItems")
            Dim pinnedItems As IEnumerable(Of PinnedItem) = collection.Find(Function(f) f.Pidl.Equals(pidl.ToString()))
            If pinnedItems.Count = 1 Then
                e = New PinnedItemEventArgs() With {.Pidl = pidl, .Index = pinnedItems(0).Index}
                collection.Delete(pinnedItems(0).Id)

                ' move items up
                For Each movePinnedItem In collection.Query().Where(Function(i) i.Index > e.Index).ToList()
                    movePinnedItem.Index -= 1
                    collection.Update(movePinnedItem)
                Next
            End If
        End Using

        If Not e Is Nothing Then
            RaiseEvent ItemUnpinned(Nothing, e)
            Return e.Index
        Else
            Return -1
        End If
    End Function

    Public Shared Sub RenameItem(oldPidl As Pidl, newPidl As Pidl)
        Using db = New LiteDatabase(getDBFileName())
            ' update in db
            Dim collection As ILiteCollection(Of PinnedItem) = db.GetCollection(Of PinnedItem)("PinnedItems")
            Dim pinnedItems As IEnumerable(Of PinnedItem) = collection.Find(Function(f) f.Pidl.Equals(oldPidl.ToString()))
            If pinnedItems.Count = 1 Then
                Dim pinnedItem As PinnedItem = pinnedItems(0)
                pinnedItem.Pidl = newPidl.ToString()
                collection.Update(pinnedItem)
            End If
        End Using
    End Sub

    Private Shared Function getDBFileName() As String
        Dim path As String = IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Laila", "Shell")
        If Not IO.Directory.Exists(path) Then IO.Directory.CreateDirectory(path)
        Return IO.Path.Combine(path, "PinnedItems3.db")
    End Function

    Public Class PinnedItem
        Public Property Id As Integer
        Public Property Pidl As String
        Public Property Index As Long
    End Class
End Class
