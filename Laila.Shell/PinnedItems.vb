Imports System.Threading
Imports Laila.Shell.Events
Imports LiteDB

Public Class PinnedItems
    Public Shared Event ItemPinned(sender As Object, e As PinnedItemEventArgs)
    Public Shared Event ItemUnpinned(sender As Object, e As PinnedItemEventArgs)

    Private Shared _lock As Object = New Object()

    Public Shared Property IsNotifying As Boolean = True

    Public Shared Function GetPinnedItems() As IEnumerable(Of Item)
        Dim pinnedItems As List(Of PinnedItem)

        Dim isSuccess As Boolean, numTries As Integer
        While (Not isSuccess AndAlso numTries <= 5)
            Try
                SyncLock _lock
                    Using db = New LiteDatabase(getDBFileName())
                        Dim collection As ILiteCollection(Of PinnedItem) = db.GetCollection(Of PinnedItem)("PinnedItems")

                        pinnedItems = collection.Query() _
                        .OrderBy(Function(f) f.Index).ToList()
                    End Using
                End SyncLock

                ' return existing pinned items, delete others
                Return Shell.GlobalThreadPool.Run(
                    Function() As IEnumerable(Of Item)
                        Dim existingPinnedItems As List(Of Item) = New List(Of Item)()
                        For Each pinnedItem In pinnedItems
                            Dim pidl As Pidl = Nothing
                            Try
                                pidl = New Pidl(pinnedItem.Pidl)
                                Dim i As Item = Item.FromPidl(pidl, Nothing)
                                If Not i Is Nothing AndAlso i.IsExisting Then
                                    i.CanShowInTree = True
                                    i.IsPinned = True
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

                        Return existingPinnedItems
                    End Function)
                isSuccess = True
            Catch ex As IO.IOException
                numTries += 1
                Thread.Sleep(500)
            End Try
        End While

        Return Array.Empty(Of Folder)()
    End Function

    Public Shared Function GetIsPinned(item As Item) As Boolean
        Dim isSuccess As Boolean, numTries As Integer
        While (Not isSuccess AndAlso numTries <= 5)
            Try
                SyncLock _lock
                    Using db = New LiteDatabase(getDBFileName())
                        ' register in db
                        Dim collection As ILiteCollection(Of PinnedItem) = db.GetCollection(Of PinnedItem)("PinnedItems")
                        Dim pinnedItems As IEnumerable(Of PinnedItem) = collection _
                            .Find(Function(f) Not item.Pidl Is Nothing AndAlso Not f.Pidl Is Nothing AndAlso Not f.Pidl Is Nothing AndAlso f.Pidl.Equals(item.Pidl.ToString()))
                        Return pinnedItems.Count > 0
                    End Using
                End SyncLock
                isSuccess = True
            Catch ex As IO.IOException
                numTries += 1
                Thread.Sleep(500)
            End Try
        End While

        Return False
    End Function

    Public Shared Sub PinItem(item As Item, Optional newIndex As Long = -1)
        Dim e As PinnedItemEventArgs = Nothing

        Dim isSuccess As Boolean, numTries As Integer
        While (Not isSuccess AndAlso numTries <= 5)
            Try
                SyncLock _lock
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
                            e = New PinnedItemEventArgs() With {.Pidl = item.Pidl, .Index = pinnedItem.Index}
                        End If
                    End Using
                End SyncLock

                If IsNotifying AndAlso Not e Is Nothing Then RaiseEvent ItemPinned(Nothing, e)
                isSuccess = True
            Catch ex As IO.IOException
                numTries += 1
                Thread.Sleep(500)
            End Try
        End While
    End Sub

    Public Shared Function UnpinItem(pidl As Pidl) As Integer
        Dim e As PinnedItemEventArgs = Nothing

        Dim isSuccess As Boolean, numTries As Integer
        While (Not isSuccess AndAlso numTries <= 5)
            Try
                SyncLock _lock
                    Using db = New LiteDatabase(getDBFileName())
                        ' register in db
                        Dim collection As ILiteCollection(Of PinnedItem) = db.GetCollection(Of PinnedItem)("PinnedItems")
                        Dim pinnedItems As IEnumerable(Of PinnedItem) = collection.Find(Function(f) f.Pidl.Equals(pidl.ToString()))
                        If pinnedItems.Count > 0 Then
                            e = New PinnedItemEventArgs() With {.Pidl = pidl, .Index = pinnedItems(0).Index}
                            collection.Delete(pinnedItems(0).Id)

                            ' move items up
                            For Each movePinnedItem In collection.Query().Where(Function(i) i.Index > e.Index).ToList()
                                movePinnedItem.Index -= 1
                                collection.Update(movePinnedItem)
                            Next
                        End If
                    End Using
                End SyncLock

                If Not e Is Nothing Then
                    If IsNotifying Then RaiseEvent ItemUnpinned(Nothing, e)
                    Return e.Index
                Else
                    Return -1
                End If
                isSuccess = True
            Catch ex As IO.IOException
                numTries += 1
                Thread.Sleep(500)
            End Try
        End While

        Return -1
    End Function

    Public Shared Sub NotifyReset()
        Dim e As PinnedItemEventArgs = New PinnedItemEventArgs() With {.Index = -1}
        RaiseEvent ItemUnpinned(Nothing, e)
    End Sub

    Public Shared Sub RenameItem(oldPidl As Pidl, newPidl As Pidl)
        Dim isSuccess As Boolean, numTries As Integer
        While (Not isSuccess AndAlso numTries <= 5)
            Try
                SyncLock _lock
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
                End SyncLock
                isSuccess = True
            Catch ex As IO.IOException
                numTries += 1
                Thread.Sleep(500)
            End Try
        End While
    End Sub

    Private Shared Function getDBFileName() As String
        Dim path As String = IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Laila", "Shell")
        If Not IO.Directory.Exists(path) Then IO.Directory.CreateDirectory(path)
        Return IO.Path.Combine(path, "PinnedItems4.db")
    End Function

    Public Class PinnedItem
        Public Property Id As Integer
        Public Property Pidl As String
        Public Property Index As Long
    End Class
End Class
