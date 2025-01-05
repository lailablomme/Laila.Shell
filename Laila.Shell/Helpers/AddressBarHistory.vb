Imports Laila.Shell.FrequentFolders
Imports LiteDB

Namespace Helpers
    Public Class AddressBarHistory
        Public Shared Function GetHistory(searchTerm As String) As IEnumerable(Of Item)
            Using db = New LiteDatabase(getDBFileName())
                Dim collection As ILiteCollection(Of AddressBarHistoryItem) = db.GetCollection(Of AddressBarHistoryItem)("AddressBarHistoryItems")

                ' return history
                Dim history1 As List(Of AddressBarHistoryItem) = collection.Query() _
                    .Where(Function(i) i.FileName.Contains(searchTerm.ToLower()) OrElse i.DisplayName.Contains(searchTerm.ToLower())) _
                    .OrderByDescending(Function(f) f.LastAccessedDateTime) _
                    .Limit(25).ToList()

                Dim tcs As New TaskCompletionSource(Of IEnumerable(Of Item))

                Shell.STATaskQueue.Add(
                    Sub()
                        Try
                            Dim history2 As List(Of Item) = New List(Of Item)()
                            Dim count As Integer = 0
                            For Each historyItem In history1
                                Dim pidl As Pidl = New Pidl(historyItem.Pidl)
                                Dim i As Item = Item.FromPidl(pidl.AbsolutePIDL, Nothing)
                                If Not i Is Nothing Then
                                    history2.Add(i)
                                    count += 1
                                    If count = 15 Then Exit For
                                End If
                                pidl.Dispose()
                            Next

                            tcs.SetResult(history2)
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
        End Function

        Public Shared Sub Track(folder As Item)
            Using db = New LiteDatabase(getDBFileName())
                ' register in db
                Dim collection As ILiteCollection(Of AddressBarHistoryItem) = db.GetCollection(Of AddressBarHistoryItem)("AddressBarHistoryItems")
                Dim items As IEnumerable(Of AddressBarHistoryItem) = collection.Find(Function(f) f.Pidl.Equals(BitConverter.ToString(folder.Pidl.Bytes)))
                If items.Count = 1 Then
                    Dim historyItem As AddressBarHistoryItem = items(0)
                    historyItem.LastAccessedDateTime = DateTime.Now
                    collection.Update(historyItem)
                Else
                    Dim historyItem As AddressBarHistoryItem = New AddressBarHistoryItem() With {
                        .Pidl = BitConverter.ToString(folder.Pidl.Bytes),
                        .FileName = If(folder.FullPath.Trim(IO.Path.DirectorySeparatorChar).LastIndexOf(IO.Path.DirectorySeparatorChar) <> -1,
                                       folder.FullPath.Trim(IO.Path.DirectorySeparatorChar).Substring(folder.FullPath.Trim(IO.Path.DirectorySeparatorChar).LastIndexOf(IO.Path.DirectorySeparatorChar) + 1),
                                       folder.FullPath.Trim(IO.Path.DirectorySeparatorChar)).ToLower(),
                        .DisplayName = folder.DisplayName.ToLower(),
                        .LastAccessedDateTime = DateTime.Now
                    }
                    collection.Insert(historyItem)
                    collection.EnsureIndex(Of String)(Function(f) f.Pidl)
                End If
            End Using
        End Sub

        Private Shared Function getDBFileName() As String
            Dim path As String = IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Laila", "Shell")
            If Not IO.Directory.Exists(path) Then IO.Directory.CreateDirectory(path)
            Return IO.Path.Combine(path, "AddressBarHistory2.db")
        End Function

        Public Class AddressBarHistoryItem
            Public Property Id As Integer
            Public Property Pidl As String
            Public Property FileName As String
            Public Property DisplayName As String
            Public Property LastAccessedDateTime As DateTime
        End Class
    End Class
End Namespace