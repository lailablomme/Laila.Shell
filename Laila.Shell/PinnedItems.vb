Imports System.IO
Imports Laila.Shell.Events
Imports Laila.Shell.FrequentFolders
Imports LiteDB

Public Class PinnedItems
    Public Shared Event ItemPinned(sender As Object, e As PinnedItemEventArgs)
    Public Shared Event ItemUnpinned(sender As Object, e As PinnedItemEventArgs)

    Public Shared Function GetPinnedItems() As IEnumerable(Of Item)
        Using db = New LiteDatabase(getDBFileName())
            Dim collection As ILiteCollection(Of PinnedItem) = db.GetCollection(Of PinnedItem)("PinnedItems")

            ' return pinned items
            Return collection.Query() _
                .OrderBy(Function(f) f.Index).ToList() _
                .Select(Function(f) Item.FromParsingNameDeepGetReverse(f.FullPath)) _
                .Where(Function(f) Not f Is Nothing) _
                .Cast(Of Item).ToList()
        End Using
    End Function

    Public Shared Function GetIsPinned(item As Item) As Boolean
        Using db = New LiteDatabase(getDBFileName())
            ' register in db
            Dim collection As ILiteCollection(Of PinnedItem) = db.GetCollection(Of PinnedItem)("PinnedItems")
            Dim pinnedItems As IEnumerable(Of PinnedItem) = collection.Find(Function(f) f.FullPath.ToLower().Equals(item.FullPath.ToLower()))
            Return pinnedItems.Count = 1
        End Using
    End Function

    Public Shared Sub PinItem(item As Item)
        Dim e As PinnedItemEventArgs

        Using db = New LiteDatabase(getDBFileName())
            ' register in db
            Dim collection As ILiteCollection(Of PinnedItem) = db.GetCollection(Of PinnedItem)("PinnedItems")
            Dim pinnedItems As IEnumerable(Of PinnedItem) = collection.Find(Function(f) f.FullPath.ToLower().Equals(item.FullPath.ToLower()))
            If pinnedItems.Count = 0 Then
                Dim pinnedItem As PinnedItem = New PinnedItem() With {
                    .FullPath = item.FullPath,
                    .Index = If(collection.Query().Count > 0, collection.Query().OrderByDescending(Function(i) i.Index).First().Index + 1, 0)
                }
                collection.Insert(pinnedItem)
                collection.EnsureIndex(Of String)(Function(f) f.FullPath)
                e = New PinnedItemEventArgs() With {.Item = item, .Index = pinnedItem.Index}
            End If
        End Using

        If Not e Is Nothing Then RaiseEvent ItemPinned(Nothing, e)
    End Sub

    Public Shared Sub UnpinItem(item As Item)
        Dim e As PinnedItemEventArgs

        Using db = New LiteDatabase(getDBFileName())
            ' register in db
            Dim collection As ILiteCollection(Of PinnedItem) = db.GetCollection(Of PinnedItem)("PinnedItems")
            Dim pinnedItems As IEnumerable(Of PinnedItem) = collection.Find(Function(f) f.FullPath.ToLower().Equals(item.FullPath.ToLower()))
            If pinnedItems.Count = 1 Then
                e = New PinnedItemEventArgs() With {.Item = item, .Index = pinnedItems(0).Index}
                collection.Delete(pinnedItems(0).Id)
            End If
        End Using

        If Not e Is Nothing Then RaiseEvent ItemUnpinned(Nothing, e)
    End Sub

    Public Shared Sub MoveItem(item As Item, newIndex As Integer)
        Using db = New LiteDatabase(getDBFileName())
            ' register in db
            Dim collection As ILiteCollection(Of PinnedItem) = db.GetCollection(Of PinnedItem)("PinnedItems")
            Dim pinnedItems As IEnumerable(Of PinnedItem) = collection.Find(Function(f) f.FullPath.ToLower().Equals(item.FullPath.ToLower()))
            If pinnedItems.Count = 1 Then
                Dim pinnedItem As PinnedItem = pinnedItems(0)
                If newIndex > pinnedItem.Index Then
                    pinnedItems = collection.Find(Function(f) f.Index <= newIndex AndAlso f.Index > pinnedItem.Index)
                    For Each pinnedItem2 In pinnedItems
                        pinnedItem2.Index -= 1
                        collection.Update(pinnedItem2)
                    Next
                ElseIf newIndex < pinnedItem.Index Then
                    pinnedItems = collection.Find(Function(f) f.Index >= newIndex AndAlso f.Index < pinnedItem.Index)
                    For Each pinnedItem2 In pinnedItems
                        pinnedItem2.Index += 1
                        collection.Update(pinnedItem2)
                    Next
                End If
                pinnedItem.Index = newIndex
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
