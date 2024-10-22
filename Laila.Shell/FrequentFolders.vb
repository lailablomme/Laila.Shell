Imports System.IO
Imports LiteDB

Public Class FrequentFolders
    Public Shared Sub Track(folder As Folder)
        ' register with os
        Functions.SHAddToRecentDocs(SHARD.SHARD_PATHW, folder.FullPath)

        Using db = New LiteDatabase(getDBFileName())
            ' register in db
            Dim collection As ILiteCollection(Of FrequentFolder) = db.GetCollection(Of FrequentFolder)("FrequentFolders")
            Dim frequentFolders As IEnumerable(Of FrequentFolder) = collection.Find(Function(f) f.FullPath.ToLower().Equals(folder.FullPath.ToLower()))
            If frequentFolders.Count = 1 Then
                Dim frequentFolder As FrequentFolder = frequentFolders(0)
                frequentFolder.AccessCount += 1
                frequentFolder.LastAccessedDateTime = DateTime.Now
                collection.Update(frequentFolder)
            Else
                Dim frequentFolder As FrequentFolder = New FrequentFolder() With {
                    .FullPath = folder.FullPath,
                    .AccessCount = 1,
                    .LastAccessedDateTime = DateTime.Now
                }
                collection.Insert(frequentFolder)
                collection.EnsureIndex(Of String)(Function(f) f.FullPath)
            End If
        End Using
    End Sub

    Public Shared Sub RecordTimeSpent(folder As Folder, minutes As Long)
        Using db = New LiteDatabase(getDBFileName())
            ' register in db
            Dim collection As ILiteCollection(Of FrequentFolder) = db.GetCollection(Of FrequentFolder)("FrequentFolders")
            Dim frequentFolders As IEnumerable(Of FrequentFolder) = collection.Find(Function(f) f.FullPath.ToLower().Equals(folder.FullPath.ToLower()))
            If frequentFolders.Count = 1 Then
                Dim frequentFolder As FrequentFolder = frequentFolders(0)
                frequentFolder.MinutesSpent += minutes
                frequentFolder.LastAccessedDateTime = DateTime.Now
                collection.Update(frequentFolder)
            Else
                Dim frequentFolder As FrequentFolder = New FrequentFolder() With {
                    .FullPath = folder.FullPath,
                    .MinutesSpent = minutes,
                    .LastAccessedDateTime = DateTime.Now
                }
                collection.Insert(frequentFolder)
                collection.EnsureIndex(Of String)(Function(f) f.FullPath)
            End If
        End Using
    End Sub

    Public Shared Function GetMostFrequent() As IEnumerable(Of Folder)
        Using db = New LiteDatabase(getDBFileName())
            Dim collection As ILiteCollection(Of FrequentFolder) = db.GetCollection(Of FrequentFolder)("FrequentFolders")

            ' only keep folders accessed in the last 10 days
            Dim foldersToDelete As List(Of FrequentFolder) = collection.Query().ToList()
            foldersToDelete = foldersToDelete _
                .OrderByDescending(Function(f) f.LastAccessedDateTime) _
                .Where(Function(f) (DateTime.Now - f.LastAccessedDateTime).TotalDays > 10) _
                .Skip(100).ToList()
            For Each deleteFolder In foldersToDelete
                collection.Delete(deleteFolder.Id)
            Next

            ' return most frequent
            Dim mostFrequent1 As List(Of FrequentFolder) = collection.Query() _
                .OrderByDescending(Function(f) f.MinutesSpent).ToList() _
                .Where(Function(f) IO.Directory.Exists(f.FullPath)).ToList()

            Dim mostFrequent2 As List(Of Folder) = New List(Of Folder)()
            Dim count As Integer = 0
            For Each folder In mostFrequent1
                Dim f As Folder = Item.FromParsingNameDeepGet(folder.FullPath)
                If Not f Is Nothing AndAlso Not PinnedItems.GetIsPinned(f) Then
                    mostFrequent2.Add(f)
                    count += 1
                    If count = 5 Then Exit For
                ElseIf Not f Is Nothing Then
                    f.Dispose()
                End If
            Next

            Return mostFrequent2
        End Using
    End Function

    Private Shared Function getDBFileName() As String
        Dim path As String = IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Laila", "Shell")
        If Not IO.Directory.Exists(path) Then IO.Directory.CreateDirectory(path)
        Return IO.Path.Combine(path, "FrequentFolders.db")
    End Function

    Public Class FrequentFolder
        Public Property Id As Integer
        Public Property FullPath As String
        Public Property AccessCount As Long
        Public Property LastAccessedDateTime As DateTime
        Public Property MinutesSpent As Long
    End Class
End Class
