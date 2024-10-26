Imports System.IO
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
        Dim existingPinnedItems As List(Of Item) = New List(Of Item)()
        For Each pinnedItem In pinnedItems
            Dim item As Item = Item.FromParsingNameDeepGetReverse(pinnedItem.FullPath)
            If Not item Is Nothing Then
                existingPinnedItems.Add(item)
            Else
                UnpinItem(pinnedItem.FullPath)
            End If
        Next

        Return existingPinnedItems
    End Function

    Public Shared Function GetIsPinned(fullPath As String) As Boolean
        Using db = New LiteDatabase(getDBFileName())
            ' register in db
            Dim collection As ILiteCollection(Of PinnedItem) = db.GetCollection(Of PinnedItem)("PinnedItems")
            Dim pinnedItems As IEnumerable(Of PinnedItem) = collection.Find(Function(f) f.FullPath.ToLower().Equals(fullPath.ToLower()))
            Return pinnedItems.Count = 1
        End Using
    End Function

    Public Shared Sub PinItem(fullPath As String, Optional newIndex As Integer = -1)
        Dim e As PinnedItemEventArgs

        Using db = New LiteDatabase(getDBFileName())
            ' register in db
            Dim collection As ILiteCollection(Of PinnedItem) = db.GetCollection(Of PinnedItem)("PinnedItems")
            Dim pinnedItems As IEnumerable(Of PinnedItem) = collection.Find(Function(f) f.FullPath.ToLower().Equals(fullPath.ToLower()))
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
                    .FullPath = fullPath,
                    .Index = If(newIndex >= 0, newIndex, If(collection.Query().Count > 0,
                        collection.Query().OrderByDescending(Function(i) i.Index).First().Index + 1, 0))
                }
                collection.Insert(pinnedItem)
                collection.EnsureIndex(Of String)(Function(f) f.FullPath)
                e = New PinnedItemEventArgs() With {.FullPath = fullPath, .Index = pinnedItem.Index}
            End If
        End Using

        If Not e Is Nothing Then RaiseEvent ItemPinned(Nothing, e)
    End Sub

    Public Shared Function UnpinItem(fullPath As String) As Integer
        Dim e As PinnedItemEventArgs

        Using db = New LiteDatabase(getDBFileName())
            ' register in db
            Dim collection As ILiteCollection(Of PinnedItem) = db.GetCollection(Of PinnedItem)("PinnedItems")
            Dim pinnedItems As IEnumerable(Of PinnedItem) = collection.Find(Function(f) f.FullPath.ToLower().Equals(fullPath.ToLower()))
            If pinnedItems.Count = 1 Then
                e = New PinnedItemEventArgs() With {.FullPath = fullPath, .Index = pinnedItems(0).Index}
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

    Public Shared Sub RenameItem(oldFullPath As String, newFullPath As String)
        Using db = New LiteDatabase(getDBFileName())
            ' update in db
            Dim collection As ILiteCollection(Of PinnedItem) = db.GetCollection(Of PinnedItem)("PinnedItems")
            Dim pinnedItems As IEnumerable(Of PinnedItem) = collection.Find(Function(f) f.FullPath.ToLower().Equals(oldFullPath.ToLower()))
            If pinnedItems.Count = 1 Then
                Dim pinnedItem As PinnedItem = pinnedItems(0)
                pinnedItem.FullPath = newFullPath
                collection.Update(pinnedItem)
            End If
        End Using
    End Sub

    Private Shared Function getDBFileName() As String
        Dim path As String = IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Laila", "Shell")
        If Not IO.Directory.Exists(path) Then IO.Directory.CreateDirectory(path)
        Return IO.Path.Combine(path, "PinnedItems.db")
    End Function

    Public Class PinnedItem
        Public Property Id As Integer
        Public Property FullPath As String
        Public Property Index As Integer
    End Class
End Class
