Imports System.IO
Imports System.Threading
Imports LiteDB

Public Class FrequentFolders
    Private Shared _lock As SemaphoreSlim = New SemaphoreSlim(1, 1)

    Public Shared Function GetMostFrequent() As IEnumerable(Of Folder)
        _lock.Wait()
        Try
            Using db = New LiteDatabase(getDBFileName())
                Dim collection As ILiteCollection(Of FrequentFolder) = db.GetCollection(Of FrequentFolder)("FrequentFolders")

                ' only keep folders accessed in the last 30 days
                Dim foldersToDelete As List(Of FrequentFolder) = collection.Query().ToList()
                foldersToDelete = foldersToDelete _
                    .OrderByDescending(Function(f) f.LastAccessedDateTime) _
                    .Where(Function(f) (DateTime.Now - f.LastAccessedDateTime).TotalDays > 30) _
                    .Skip(100).ToList()
                For Each deleteFolder In foldersToDelete
                    collection.Delete(deleteFolder.Id)
                Next

                ' return most frequent
                Dim mostFrequent1 As List(Of FrequentFolder) = collection.Query() _
                    .OrderByDescending(Function(f) f.AccessCount).ToList()

                Return Shell.GlobalThreadPool.Run(
                    Function() As IEnumerable(Of Item)
                        Dim mostFrequent2 As List(Of Folder) = New List(Of Folder)()
                        Dim count As Integer = 0
                        For Each folder In mostFrequent1
                            Dim pidl As Pidl = Nothing
                            Try
                                pidl = New Pidl(folder.Pidl)
                                Dim i As Item = Item.FromPidl(pidl, Nothing)
                                If Not i Is Nothing AndAlso TypeOf i Is Folder AndAlso Not PinnedItems.GetIsPinned(i) Then
                                    mostFrequent2.Add(i)
                                    count += 1
                                    If count = 5 Then Exit For
                                ElseIf Not i Is Nothing Then
                                    i.Dispose()
                                End If
                            Finally
                                If Not pidl Is Nothing Then
                                    pidl.Dispose()
                                End If
                            End Try
                        Next

                        Return mostFrequent2
                    End Function)
            End Using
        Finally
            _lock.Release()
        End Try
    End Function

    Public Shared Sub Track(folder As Folder)
        If folder.Pidl Is Nothing Then
            Return
        End If

        ' register with os
        Functions.SHAddToRecentDocs(SHARD.SHARD_PATHW, folder.FullPath)

        _lock.Wait()
        Try
            Using db = New LiteDatabase(getDBFileName())
                ' register in db
                Dim collection As ILiteCollection(Of FrequentFolder) = db.GetCollection(Of FrequentFolder)("FrequentFolders")
                Dim frequentFolders As IEnumerable(Of FrequentFolder) = collection.Find(Function(f) f.Pidl.Equals(folder.Pidl.ToString()))
                If frequentFolders.Count = 1 Then
                    Dim frequentFolder As FrequentFolder = frequentFolders(0)
                    frequentFolder.AccessCount += 1
                    frequentFolder.LastAccessedDateTime = DateTime.Now
                    collection.Update(frequentFolder)
                Else
                    Dim frequentFolder As FrequentFolder = New FrequentFolder() With {
                        .Pidl = folder.Pidl.ToString(),
                        .AccessCount = 1,
                        .LastAccessedDateTime = DateTime.Now
                    }
                    collection.Insert(frequentFolder)
                End If
            End Using
        Finally
            _lock.Release()
        End Try
    End Sub

    Public Shared Sub RenameItem(oldPidl As Pidl, newPidl As Pidl)
        _lock.Wait()
        Try
            Using db = New LiteDatabase(getDBFileName())
                ' update in db
                Dim collection As ILiteCollection(Of FrequentFolder) = db.GetCollection(Of FrequentFolder)("FrequentFolders")
                Dim frequentFolders As IEnumerable(Of FrequentFolder) = collection.Find(Function(f) f.Pidl.Equals(oldPidl.ToString()))
                If frequentFolders.Count = 1 Then
                    Dim frequentFolder As FrequentFolder = frequentFolders(0)
                    frequentFolder.Pidl = newPidl.ToString()
                    collection.Update(frequentFolder)
                End If
            End Using
        Finally
            _lock.Release()
        End Try
    End Sub

    Private Shared Function getDBFileName() As String
        Dim path As String = IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Laila", "Shell")
        If Not IO.Directory.Exists(path) Then IO.Directory.CreateDirectory(path)
        Return IO.Path.Combine(path, "FrequentFolders3.db")
    End Function

    Public Class FrequentFolder
        Public Property Id As Integer
        Public Property Pidl As String
        Public Property AccessCount As Long
        Public Property LastAccessedDateTime As DateTime
    End Class
End Class
