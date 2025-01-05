Imports Laila.Shell.FrequentFolders
Imports LiteDB

Namespace Helpers
    Public Class AddressBarHistory
        Public Shared Function GetHistory() As IEnumerable(Of Folder)
            Using db = New LiteDatabase(getDBFileName())
                Dim collection As ILiteCollection(Of AddressBarHistoryFolder) = db.GetCollection(Of AddressBarHistoryFolder)("AddressBarHistoryFolders")

                ' return history
                Dim history1 As List(Of AddressBarHistoryFolder) = collection.Query() _
                .OrderByDescending(Function(f) f.LastAccessedDateTime) _
                .Limit(25).ToList()

                Dim tcs As New TaskCompletionSource(Of IEnumerable(Of Item))

                Shell.STATaskQueue.Add(
                Sub()
                    Try
                        Dim history2 As List(Of Folder) = New List(Of Folder)()
                        Dim count As Integer = 0
                        For Each folder In history1
                            Dim pidl As Pidl = New Pidl(folder.Pidl)
                            Dim f As Folder = Item.FromPidl(pidl.AbsolutePIDL, Nothing)
                            If Not f Is Nothing Then
                                history2.Add(f)
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

        Public Shared Sub Track(folder As Folder)
            Using db = New LiteDatabase(getDBFileName())
                ' register in db
                Dim collection As ILiteCollection(Of AddressBarHistoryFolder) = db.GetCollection(Of AddressBarHistoryFolder)("AddressBarHistoryFolders")
                Dim folders As IEnumerable(Of AddressBarHistoryFolder) = collection.Find(Function(f) f.Pidl.Equals(BitConverter.ToString(folder.Pidl.Bytes)))
                If folders.Count = 1 Then
                    Dim historyFolder As AddressBarHistoryFolder = folders(0)
                    historyFolder.LastAccessedDateTime = DateTime.Now
                    collection.Update(historyFolder)
                Else
                    Dim historyFolder As AddressBarHistoryFolder = New AddressBarHistoryFolder() With {
                        .Pidl = BitConverter.ToString(folder.Pidl.Bytes),
                        .LastAccessedDateTime = DateTime.Now
                    }
                    collection.Insert(historyFolder)
                    collection.EnsureIndex(Of String)(Function(f) f.Pidl)
                End If
            End Using
        End Sub

        Private Shared Function getDBFileName() As String
            Dim path As String = IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Laila", "Shell")
            If Not IO.Directory.Exists(path) Then IO.Directory.CreateDirectory(path)
            Return IO.Path.Combine(path, "AddressBarHistory.db")
        End Function

        Public Class AddressBarHistoryFolder
            Public Property Id As Integer
            Public Property Pidl As String
            Public Property LastAccessedDateTime As DateTime
        End Class
    End Class
End Namespace