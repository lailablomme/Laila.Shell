Imports System.IO
Imports LiteDB

Public Class FrequentFolders
    Private Shared _lock As Object = New Object()

    Public Shared Function GetMostFrequent() As IEnumerable(Of Folder)
        SyncLock _lock
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
                    .OrderByDescending(Function(f) f.AccessCount).ToList() _
                    .OrderByDescending(Function(f) f.MinutesSpent).ToList() _
                    .Where(Function(f) Not String.IsNullOrWhiteSpace(f.FullPath) _
                        AndAlso Not f.FullPath.StartsWith("\\") AndAlso IO.Directory.Exists(f.FullPath)).ToList()

                Dim tcs As New TaskCompletionSource(Of IEnumerable(Of Item))

                Shell.STATaskQueue.Add(
                    Sub()
                        Try
                            Dim mostFrequent2 As List(Of Folder) = New List(Of Folder)()
                            Dim count As Integer = 0
                            For Each folder In mostFrequent1
                                Dim f As Folder = Item.FromParsingName(folder.FullPath, Nothing, True)
                                If Not f Is Nothing AndAlso Not PinnedItems.GetIsPinned(f.FullPath) Then
                                    mostFrequent2.Add(f)
                                    count += 1
                                    If count = 5 Then Exit For
                                ElseIf Not f Is Nothing Then
                                    f.Dispose()
                                End If
                            Next

                            tcs.SetResult(mostFrequent2)
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
            End Using
        End SyncLock
    End Function

    Public Shared Sub Track(folder As Folder)
        ' register with os
        Functions.SHAddToRecentDocs(SHARD.SHARD_PATHW, folder.FullPath)

        SyncLock _lock
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
        End SyncLock
    End Sub

    Public Shared Sub RecordTimeSpent(folder As Folder, minutes As Long)
        SyncLock _lock
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
        End SyncLock
    End Sub

    Public Shared Sub RenameItem(oldFullPath As String, newFullPath As String)
        SyncLock _lock
            Using db = New LiteDatabase(getDBFileName())
                ' update in db
                Dim collection As ILiteCollection(Of FrequentFolder) = db.GetCollection(Of FrequentFolder)("FrequentFolders")
                Dim frequentFolders As IEnumerable(Of FrequentFolder) = collection.Find(Function(f) f.FullPath.ToLower().Equals(oldFullPath.ToLower()))
                If frequentFolders.Count = 1 Then
                    Dim frequentFolder As FrequentFolder = frequentFolders(0)
                    frequentFolder.FullPath = newFullPath
                    collection.Update(frequentFolder)
                End If
            End Using
        End SyncLock
    End Sub

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
